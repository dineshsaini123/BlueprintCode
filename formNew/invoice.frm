VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form invoice 
   ClientHeight    =   10248
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10248
   ScaleWidth      =   14880
   Begin VB.Frame panel 
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   10155
      Left            =   72
      TabIndex        =   24
      Top             =   90
      Width           =   14700
      Begin VB.CommandButton cmdRef 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Order No"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   3456
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   9612
         Width           =   1236
      End
      Begin VB.CommandButton cmdMaster 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Master Details"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   2016
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   9612
         Width           =   1416
      End
      Begin VB.CheckBox Check1_disZero 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   432
         Left            =   1116
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   9612
         Width           =   876
      End
      Begin VB.TextBox txtchecked 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   3468
         MaxLength       =   100
         TabIndex        =   112
         Top             =   8424
         Width           =   540
      End
      Begin VB.CommandButton cmdListBlankOrd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Empty No"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   108
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   9612
         Width           =   948
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
         Height          =   7224
         Left            =   288
         TabIndex        =   110
         Top             =   2376
         Visible         =   0   'False
         Width           =   1884
      End
      Begin VB.CheckBox Check1_spremarks 
         Caption         =   "Edit Specimen Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   11475
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   8370
         Width           =   2670
      End
      Begin VB.CheckBox Check1_withheader 
         Caption         =   "With Header"
         Height          =   195
         Left            =   7908
         TabIndex        =   105
         Top             =   9672
         Width           =   1212
      End
      Begin VB.TextBox txtPendingBooksRem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8145
         MaxLength       =   100
         TabIndex        =   86
         Top             =   8010
         Width           =   5985
      End
      Begin VB.CheckBox Check1_notPrint_inst 
         Caption         =   "Not Print Instruction"
         Height          =   195
         Left            =   9585
         TabIndex        =   101
         Top             =   8505
         Width           =   1800
      End
      Begin VB.TextBox txtappno 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   4992
         MaxLength       =   100
         TabIndex        =   100
         Top             =   8460
         Width           =   540
      End
      Begin VB.ComboBox txtPlaceSup 
         Height          =   315
         Left            =   4995
         TabIndex        =   18
         Top             =   2565
         Width           =   2520
      End
      Begin VB.CheckBox Check1_trans 
         Caption         =   "Transport Copy"
         Height          =   195
         Left            =   5685
         TabIndex        =   92
         Top             =   8520
         Width           =   1395
      End
      Begin VB.CheckBox Check1_dos 
         Caption         =   "Show Screen  "
         Height          =   195
         Left            =   8190
         TabIndex        =   91
         Top             =   8520
         Width           =   1350
      End
      Begin VB.TextBox txtTODNO 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   675
         MaxLength       =   100
         TabIndex        =   88
         Top             =   8415
         Width           =   900
      End
      Begin VB.CheckBox option_withHeader 
         Caption         =   "Send Mail   "
         Height          =   195
         Left            =   7110
         TabIndex        =   87
         Top             =   8520
         Width           =   1125
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   810
         MaxLength       =   100
         TabIndex        =   85
         Top             =   8010
         Width           =   5850
      End
      Begin VB.CheckBox Check1_school 
         Caption         =   "Select School"
         Height          =   195
         Left            =   12555
         TabIndex        =   83
         Top             =   2610
         Width           =   1335
      End
      Begin VB.TextBox txtShip 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5220
         TabIndex        =   15
         Top             =   1740
         Width           =   8835
      End
      Begin VB.TextBox txtAmtwords 
         Appearance      =   0  'Flat
         Height          =   540
         Left            =   9615
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   9045
         Width           =   4584
      End
      Begin VB.TextBox txtScId 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   13500
         TabIndex        =   77
         Top             =   1395
         Width           =   555
      End
      Begin VB.CommandButton cmdPendingClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete Order From Pending List"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   7215
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton cmdPendingOrder 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pending Order"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   6195
         Picture         =   "invoice.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   420
         Width           =   996
      End
      Begin VB.ComboBox customercode 
         Appearance      =   0  'Flat
         Height          =   912
         Left            =   9348
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Top             =   405
         Visible         =   0   'False
         Width           =   4725
      End
      Begin VB.ComboBox txtMark 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "invoice.frx":0BE4
         Left            =   7185
         List            =   "invoice.frx":0BF1
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1410
         Width           =   795
      End
      Begin VB.ComboBox cmbtransportname 
         Height          =   315
         Left            =   2205
         TabIndex        =   17
         Top             =   2565
         Width           =   2790
      End
      Begin VB.TextBox txtadst 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1905
         TabIndex        =   41
         Top             =   7620
         Width           =   1035
      End
      Begin VB.ComboBox cmbAgentName 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "invoice.frx":0BFE
         Left            =   9345
         List            =   "invoice.frx":0C00
         TabIndex        =   8
         Top             =   765
         Width           =   4725
      End
      Begin VB.CommandButton Commandall 
         Caption         =   "All Books"
         Height          =   420
         Left            =   -315
         TabIndex        =   40
         Top             =   6075
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Commandother 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&End Part"
         Enabled         =   0   'False
         Height          =   435
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   7485
         Width           =   930
      End
      Begin VB.ComboBox Bookname 
         Height          =   2064
         Left            =   4896
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   38
         Top             =   3312
         Width           =   2355
      End
      Begin VB.ComboBox Bookcode 
         Height          =   2256
         ItemData        =   "invoice.frx":0C02
         Left            =   2700
         List            =   "invoice.frx":0C04
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   37
         Top             =   3090
         Width           =   2355
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00B8E4F1&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   195
         ScaleHeight     =   756
         ScaleWidth      =   9192
         TabIndex        =   26
         Top             =   8820
         Width           =   9195
         Begin VB.CommandButton CommandDirectPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Direct Print"
            Enabled         =   0   'False
            Height          =   645
            Left            =   7110
            Picture         =   "invoice.frx":0C06
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   45
            Width           =   990
         End
         Begin VB.CommandButton Commandadd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   645
            Left            =   15
            Picture         =   "invoice.frx":17EA
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   45
            Width           =   990
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   645
            Left            =   8145
            Picture         =   "invoice.frx":23CE
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   45
            Width           =   990
         End
         Begin VB.CommandButton CommandPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Enabled         =   0   'False
            Height          =   645
            Left            =   6075
            Picture         =   "invoice.frx":2FB2
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   45
            Width           =   990
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   645
            Left            =   5025
            Picture         =   "invoice.frx":3B96
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   45
            Width           =   990
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Enabled         =   0   'False
            Height          =   645
            Left            =   4020
            Picture         =   "invoice.frx":477A
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   45
            Width           =   990
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   645
            Left            =   3045
            Picture         =   "invoice.frx":535E
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   45
            Width           =   990
         End
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Height          =   645
            Left            =   2040
            Picture         =   "invoice.frx":58E8
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   45
            Width           =   990
         End
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   645
            Left            =   1035
            Picture         =   "invoice.frx":64CC
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   45
            Width           =   990
         End
         Begin VB.CommandButton Commandhelp 
            Caption         =   "Help"
            Height          =   495
            Left            =   -720
            TabIndex        =   27
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CommandButton Commandprintnh 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N&HPrint"
            Enabled         =   0   'False
            Height          =   645
            Left            =   6420
            Picture         =   "invoice.frx":690E
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   45
            Visible         =   0   'False
            Width           =   75
         End
      End
      Begin VB.ComboBox Genledger 
         Height          =   315
         Left            =   9375
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   7590
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   4290
         Left            =   90
         TabIndex        =   23
         Top             =   2970
         Width           =   12855
         _ExtentX        =   22670
         _ExtentY        =   7557
         _Version        =   393216
         BackColorFixed  =   7917545
         ForeColorFixed  =   16711680
         BackColorBkg    =   16777215
         FillStyle       =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox textbox 
         Height          =   315
         Left            =   9345
         TabIndex        =   6
         Top             =   420
         Width           =   4725
         _ExtentX        =   8340
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox through 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   1410
         Width           =   2415
         _ExtentX        =   4255
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_DTOB 
         Height          =   315
         Left            =   4230
         TabIndex        =   4
         Top             =   780
         Width           =   990
         _ExtentX        =   1736
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox bundles 
         Height          =   285
         Left            =   1095
         TabIndex        =   10
         Top             =   1410
         Width           =   1155
         _ExtentX        =   2053
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_OB 
         Height          =   315
         Left            =   2085
         TabIndex        =   2
         Top             =   780
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox i_dt 
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
         Left            =   1035
         TabIndex        =   1
         Top             =   780
         Width           =   1050
         _ExtentX        =   1863
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tempmeb 
         Height          =   285
         Left            =   90
         TabIndex        =   42
         Top             =   2925
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1545
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rate 
         Height          =   285
         Left            =   1110
         TabIndex        =   43
         Top             =   4530
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3239
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox amount 
         Height          =   285
         Left            =   150
         TabIndex        =   44
         Top             =   4920
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3239
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_NO 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   765
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox through1 
         Height          =   285
         Left            =   4710
         TabIndex        =   12
         Top             =   1410
         Width           =   2475
         _ExtentX        =   4360
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox marka 
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   1395
         Width           =   975
         _ExtentX        =   1715
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox weight 
         Height          =   315
         Left            =   11250
         TabIndex        =   22
         Top             =   2565
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox freight 
         Height          =   315
         Left            =   10035
         TabIndex        =   21
         Top             =   2565
         Width           =   1095
         _ExtentX        =   1947
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox bdated 
         Height          =   315
         Left            =   8925
         TabIndex        =   20
         Top             =   2565
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox biltno 
         Height          =   315
         Left            =   7515
         TabIndex        =   19
         Top             =   2565
         Width           =   1365
         _ExtentX        =   2413
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox station 
         Height          =   315
         Left            =   90
         TabIndex        =   16
         Top             =   2565
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtOrderNo 
         Height          =   315
         Left            =   3315
         TabIndex        =   3
         Top             =   780
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtschool 
         Height          =   285
         Left            =   8520
         TabIndex        =   14
         Top             =   1395
         Width           =   4965
         _ExtentX        =   8763
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTODDate 
         Height          =   270
         Left            =   1620
         TabIndex        =   90
         Top             =   8415
         Width           =   990
         _ExtentX        =   1757
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNSCHNo 
         Height          =   315
         Left            =   5220
         TabIndex        =   5
         Top             =   780
         Width           =   945
         _ExtentX        =   1672
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777152
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtcreatedDT 
         Height          =   288
         Left            =   12528
         TabIndex        =   107
         Top             =   7632
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRandomDT 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   312
         Left            =   5832
         TabIndex        =   114
         Top             =   9648
         Visible         =   0   'False
         Width           =   948
         _ExtentX        =   1672
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRandomId 
         Height          =   312
         Left            =   4716
         TabIndex        =   115
         Top             =   9648
         Visible         =   0   'False
         Width           =   1104
         _ExtentX        =   1947
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRandomMob 
         Height          =   312
         Left            =   6804
         TabIndex        =   116
         Top             =   9648
         Visible         =   0   'False
         Width           =   1068
         _ExtentX        =   1884
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Checked :"
         Height          =   252
         Left            =   2700
         TabIndex        =   113
         Top             =   8460
         Width           =   684
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Postage :"
         Height          =   252
         Left            =   11808
         TabIndex        =   109
         Top             =   9648
         Width           =   660
      End
      Begin VB.Label lblPostage 
         BackColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   12492
         TabIndex        =   108
         Top             =   9648
         Width           =   600
      End
      Begin VB.Label lblPartyfrt 
         BackColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   13752
         TabIndex        =   104
         Top             =   9648
         Width           =   492
      End
      Begin VB.Label lblfrt 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Freight :"
         Height          =   252
         Left            =   13104
         TabIndex        =   103
         Top             =   9648
         Width           =   660
      End
      Begin VB.Label Label29 
         Caption         =   "Specimen Remarks"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6705
         TabIndex        =   102
         Top             =   8055
         Width           =   1485
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "App. :"
         Height          =   252
         Left            =   4368
         TabIndex        =   99
         Top             =   8496
         Width           =   432
      End
      Begin VB.Label lblApp 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   4860
         TabIndex        =   98
         Top             =   8460
         Width           =   210
      End
      Begin VB.Label lblSMSId 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   97
         Top             =   8505
         Width           =   1365
      End
      Begin VB.Label LblRandomNo 
         Height          =   240
         Left            =   13140
         TabIndex        =   96
         Top             =   2115
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NS-CH.No "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5265
         TabIndex        =   94
         Top             =   450
         Width           =   900
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Place Of Supply "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5070
         TabIndex        =   93
         Top             =   2340
         Width           =   1545
      End
      Begin VB.Label Label25 
         Caption         =   "V. No :"
         Height          =   255
         Left            =   135
         TabIndex        =   89
         Top             =   8460
         Width           =   795
      End
      Begin VB.Label Label24 
         Caption         =   "Remarks"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   135
         TabIndex        =   84
         Top             =   8040
         Width           =   1035
      End
      Begin VB.Label lblPAN 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   1164
         TabIndex        =   82
         Top             =   1800
         Width           =   1272
      End
      Begin VB.Label lblBookSId 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13695
         TabIndex        =   81
         Top             =   2055
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ship to: "
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   4680
         TabIndex        =   80
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label23 
         Caption         =   "Amt in words :"
         Height          =   255
         Left            =   9615
         TabIndex        =   79
         Top             =   8850
         Width           =   1035
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "School : "
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   7965
         TabIndex        =   76
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lblMail 
         BackColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   9252
         TabIndex        =   74
         Top             =   9648
         Width           =   2520
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OrderNo"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3270
         TabIndex        =   73
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mark "
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7305
         TabIndex        =   71
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transport"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2355
         TabIndex        =   70
         Top             =   2340
         Width           =   1935
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Representative :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8235
         TabIndex        =   69
         Top             =   825
         Width           =   1170
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F4 Key To Delete A Invoive Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   150
         TabIndex        =   68
         Top             =   7275
         Width           =   2955
      End
      Begin VB.Label labelbybanklbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "By Bank : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5955
         TabIndex        =   67
         Top             =   7650
         Width           =   795
      End
      Begin VB.Label labelbybank 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6780
         TabIndex        =   66
         Top             =   7650
         Width           =   885
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8850
         TabIndex        =   65
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Marka : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         TabIndex        =   64
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Total Quantity : "
         Height          =   255
         Left            =   3390
         TabIndex        =   63
         Top             =   5010
         Width           =   1470
      End
      Begin VB.Label tqu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4965
         TabIndex        =   62
         Top             =   7650
         Width           =   930
      End
      Begin VB.Label mgd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11250
         TabIndex        =   61
         Top             =   7350
         Width           =   1200
      End
      Begin VB.Label mna 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11250
         TabIndex        =   60
         Top             =   7650
         Width           =   1200
      End
      Begin VB.Label mga 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   59
         Top             =   7350
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   975
         TabIndex        =   58
         Top             =   405
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   75
         TabIndex        =   57
         Top             =   405
         Width           =   1050
      End
      Begin VB.Label label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Code : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8235
         TabIndex        =   56
         Top             =   465
         Width           =   1185
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Net Amount : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10050
         TabIndex        =   55
         Top             =   7650
         Width           =   1155
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Gross Amount : "
         Height          =   255
         Left            =   6750
         TabIndex        =   54
         Top             =   4860
         Width           =   1260
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Order By "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2130
         TabIndex        =   53
         Top             =   405
         Width           =   915
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4140
         TabIndex        =   52
         Top             =   405
         Width           =   900
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bundle(s) : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   51
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Through : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2370
         TabIndex        =   50
         Top             =   1170
         Width           =   5355
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Railway/Station : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         TabIndex        =   49
         Top             =   2340
         Width           =   1230
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bilty No. : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7545
         TabIndex        =   48
         Top             =   2340
         Width           =   1230
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Freight : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10170
         TabIndex        =   47
         Top             =   2295
         Width           =   780
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Weight "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   11430
         TabIndex        =   46
         Top             =   2340
         Width           =   615
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Total Discount : "
         Height          =   255
         Left            =   8070
         TabIndex        =   45
         Top             =   4800
         Width           =   1290
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   816
         Left            =   180
         Top             =   8796
         Width           =   9276
      End
   End
End
Attribute VB_Name = "invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kk As ADODB.Recordset
Dim kk1 As ADODB.Recordset
'Dim CON As ADODB.Connection
Dim I As Integer
Dim lastrow, lastcol As Integer
Dim VALIDRATE As Boolean
Dim maxrow As Integer
Public totalamount, totaldiscount As Double
Public otheramount, otherdiscount As Double
Dim autoscroll As Boolean
Public Edit As Boolean
Dim addmode As Boolean
Dim Printheader As Boolean
Dim addoredit As Boolean
Dim category As String
Dim bkdesc As String
Dim emptyInv_bool As Boolean
Sub printinvoice()

Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = False
    Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
            Me.Commanddelete.Enabled = False
        Me.Commandabandon.Enabled = True
    Me.CommandPrint.Enabled = True
Me.Commandprintnh.Enabled = True


Dim called1, called2 As Boolean
Dim MaxLine As Integer
Dim netamount As Double
Dim totalquantity As Long
Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
Dim RS As ADODB.Recordset
Dim Pno As Integer
Set RS = New ADODB.Recordset
T1 = 10
T2 = 25
T3 = 40
T4 = 55
T5 = 70
T6 = 85
T7 = 100
T8 = 115
netamount = 0
totalquantity = 0
paperWidth = 96
MaxLine = 60
called1 = False
called2 = False

Dim Line As Integer
Dim rs1 As ADODB.Recordset
Dim kkk As ADODB.Recordset
Dim FooterYes As Boolean
Set kkk = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim mystr1 As String
mystr1 = ""
If Me.txtMark.text = "M" Then
mystr1 = "MOHKAMPUR"
ElseIf Me.txtMark.text = "W" Then
mystr1 = "W.K.ROAD"
ElseIf Me.txtMark.text = "U" Then
mystr1 = "UTSAV COMPLEX"
End If

Dim LEFTM As Integer
Open "" + VB.App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
FooterYes = False
header:
    If kkk.State = 1 Then
          kkk.Close
    End If
    CNSetup
    kkk.Open "select * from setup1", con, adOpenStatic, adLockReadOnly, adCmdText
    If FooterYes = True Then
    
        If Line > MaxLine - 6 Then
            Do While Line < 60
                Print #1, ""
                Line = Line + 1
            Loop
        End If
        Line = 0
        LEFTM = 5
        Print #1, Tab(0); repli("-", 96)
        'Print #1, Tab(1); "E.& O.E"
        'Print #1, Tab(1); kkk!COURT; Tab(LEFTM + (paperWidth - ((Len(kkk!COURT) + Len(kkk!Cname)) * 0.75))); "FOR " + Trim(kkk!Cname)
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); kkk!COURT; Tab(65); "FOR " + Trim(kkk!cname)
        Print #1, ""
        Print #1, Tab(1); Chr(27) + Chr(71); "Continued on Page : " & Pno; Chr(27) + Chr(72)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
End If
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

Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("I N V O I C E")))) / 2 - 3); Chr(14); "I N V O I C E"; Chr(20); Tab(52); IIf(Printheader = True, kkk!uptt, "")
Line = Line + 1
If Printheader = True Then
   Print #1, Tab(63); kkk!cst
   Line = Line + 1
End If
If Printheader = False Then
   Print #1, ""
   Line = Line + 1
End If
Print #1, repli("-", 96)
Line = Line + 1
If rs1.State = 1 Then rs1.Close
rs1.Open "select top 10 * from invoicea where invoiceno='" + Trim(Me.I_NO.text) + "' and " & stringyear & "", con, adOpenStatic, adLockReadOnly
'rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); "To,   S.L. Code : "; Tab(20); Mid$(rs1!subledger, 1, 5); Tab(50); "Invoice No. : "; Chr(27) + Chr(72); Trim(rs1!invoiceNo); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); rs1!invoiceDate
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & kkk!DESCFORINVOICE
        Print #1, Tab(5); IIf(IsNull(kkk!address1), " ", kkk!address1); Tab(49); Chr(27) + Chr(71); "Order by    : "; Chr(27) + Chr(72); Trim(rs1!orderby); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!address2), " ", kkk!address2)
        Print #1, Tab(5); IIf(IsNull(kkk!address3), " ", kkk!address3); Tab(49); Chr(27) + Chr(71); "Bilty No.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        'Print #1, ""
        Print #1, Tab(73); Chr(27) + Chr(71); "(" & txtMark & ")"; Chr(27) + Chr(72)
        kkk.Close
        Print #1, Chr(27) + Chr(71); "Through  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!through) + IIf(Trim(rs1!through1) = "", "", "," & rs1!through1); Tab(71); Chr(27) + Chr(71); "Agent Name : "; Chr(27) + Chr(72); Trim(rs1!agentname)
        Print #1, Chr(27) + Chr(71); "Station  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!station); " "; Trim(rs1!transportname); Tab(71); Chr(27) + Chr(71); "Pvt. Mark   : "; Chr(27) + Chr(72); Trim(rs1!marka)
        Print #1, Chr(27) + Chr(71); "Freight  :"; Chr(27) + Chr(72); Tab(15); Trim(IIf(IsNull(rs1!freight), "", rs1!freight)); Tab(40); Chr(27) + Chr(71); "Weight    : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."; Tab(73); Chr(27) + Chr(71); "Bundle(s)   : "; Chr(27) + Chr(72); Trim(rs1!bundles)
        Print #1, Chr(27) + Chr(71); repli("-", 96)
        Print #1, Tab(0); "S.No."; Tab(15); "Book Description"; Tab(50); "Quantity"; Tab(62); "Rate"; Tab(74); "Amount"; Tab(86); "Net Amount"
        Print #1, repli("-", 96); Chr(27) + Chr(72)
        Line = Line + 12
    End If
    If called1 Then
        GoTo printagain1
    End If
    If called2 Then
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.Close
    kk.Open "select * from INVOICEB where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " order by printorder,sno ", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
        kk.MoveFirst
        Dim cdiscount As Double
        Dim sno As Integer
        Dim tdata As ADODB.Recordset
        Set tdata = New ADODB.Recordset
        sno = 1
        Do While Not kk.EOF
            cdiscount = kk!PRINTORDER
            Do While kk!PRINTORDER = cdiscount
                vdis = kk!discount
                tdata.Open "Select bookname from books where " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
                Print #1, Tab(0); rsets(Trim(Str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(Str(kk!QUANTITY)), 5); Tab(58); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!QUANTITY
                Line = Line + 1
                If Line > MaxLine - 3 Then
                    called1 = True
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
printagain1:
                    called1 = False
               End If
               tdata.Close
               If Not kk.EOF Then
                    sno = sno + 1
                    kk.MoveNext
               End If
                If kk.EOF Then
                    Exit Do
                End If
            Loop
            If Line > MaxLine - 6 Then
                    called2 = True
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
printagain2:
                    
                    called2 = False
                End If
                Print #1, Tab(70); repli("-", 12)
                Line = Line + 1
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from INVOICEB where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and printorder =" + Trim(Str(cdiscount)) + " group by printorder", con, adOpenStatic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(Str(tdata(0)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(vdis), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(70); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.Close
             Loop
           End If
       End If
       Print #1, repli("-", 96)
       Print #1, Tab(50); rsets(Trim(Str(totalquantity)), 7); Tab(84); rsets(Trim(Format(Str(netamount), "0.00")), 12)
       Line = Line + 2
       If kk.State = 1 Then
             kk.Close
       End If
       kk.Open "Select * from invoicec where  " & stringyear & " and  invoiceno=" + Trim(Me.I_NO.text), con, adOpenStatic, adLockReadOnly, adCmdText
       If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(60); Trim(kk!text) + " :  @  " + Trim(Format(Str(kk!rate), "0.00")) & " % "; Tab(84); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(60); Trim(kk!text) & " :"; Tab(84); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
        End If
        Print #1, Tab(84); repli("-", 12)
        Print #1, Chr(27) + Chr(71); Tab(46); "NET AMOUNT  : "; Tab(85); rsets(Trim(Format(Str(netamount), "0.00")), 12); Chr(27) + Chr(72)
        Print #1, Tab(84); repli("-", 12)
        VNetamt = netamount
        Line = Line + 3
        kk.Close
        Dim Va As Variant
        kk.Open "Select * from invoicea where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenStatic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
             If kk!txt1a <> 0 Then
                Print #1, Tab(60); kk!txt1 & "  :"; Tab(84); rsets(Trim(Format(Str(Abs(kk!txt1a)), "0.00")), 12)
                Line = Line + 1
                Va = netamount
                Va = Va + kk!txt1a
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(60); kk!txt2 & " :"; Tab(84); rsets(Trim(Format(Str(Abs(kk!txt2a)), "0.00")), 12)
                 Line = Line + 1
                 Va = netamount
                 Va = Va - kk!txt2a
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(45); "BY BANK     :"; Tab(84); rsets(Trim(Format(Str(Abs(kk!baa)), "0.00")), 12)
                 Line = Line + 1
                 Va = netamount
                 Va = netamount - kk!baa
             End If
        
            If kk!baa <> 0 Then
               Print #1, Tab(84); repli("-", 12)
               Print #1, Tab(45); "BALANCE     : "; Tab(84); rsets(Trim(Format(Str(Va), "0.00")), 12);
               Print #1, Tab(84); repli("-", 12)
               Line = Line + 3
            End If
        End If
       'PRINT THE FOOTER IN INVOICE START
        Do While Line < 61
            Print #1, ""
            Line = Line + 1
        Loop
        Print #1, Tab(0); Chr(27) + Chr(71); toword(Round(VNetamt, 2)); Chr(27) + Chr(72)
        Print #1, Tab(0); repli("-", 96)
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
        Print #1, ""
        Print #1, ""
        'PRINT THE FOOTER IN INVOICE END
        Close #1
        
        
        
        ''PrintOption.Show
        
        

End Sub
Sub invoicecalc()
'OTHERSALES.calc
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     
     
     
     
     mna.Caption = Format(Round((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
End Sub
Sub invoiceabandon()

Check1_disZero.value = 0
txtchecked.text = ""
txtRandomDT.text = "__/__/____"
txtRandomId.text = ""
txtRandomMob.text = ""

lblPostage.Caption = ""

txtcreatedDT.text = ""
Check1_school.value = 0
lblPartyfrt.Caption = ""

Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = True
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = True
Me.Commandabandon.Enabled = True
Me.CommandPrint.Enabled = True
Me.CommandDirectPrint.Enabled = True
Me.Commandprintnh.Enabled = True
        
Check1_notPrint_inst.value = 0
        
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
If kk.State = 1 Then
   kk.Close
End If

If Edit = False Then
End If
        
On Error Resume Next
Dim ctl As Control
For Each ctl In Me.Controls
    If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
        If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) Then
            ctl.text = ""
        End If
        ctl.Enabled = False
    End If
Next

For I = 1 To Grid1.rows - 1
If (Grid1.TextMatrix(I, 1) <> "" Or Grid1.TextMatrix(I, 2) <> "") Then
Grid1.Row = I
 For J = 1 To 8
     Grid1.Col = J
    Grid1.text = ""
Next
End If
Next
        
txtPendingBooksRem.text = ""
txtappno.text = ""
lblApp(4).Caption = ""
lblSMSId.Caption = ""
txtTODNO.text = ""
txtTODDate.text = "__/__/____"
       
lblBookSId.Caption = ""
lblPAN.Caption = ""
txtschool = ""
txtScId.text = ""
txtShip = ""
lblShipAdd.Caption = ""

I_DTOB = "__/__/____"
bdated = "__/__/____"
tqu.Caption = ""
mga.Caption = ""
mgd.Caption = ""
mna.Caption = ""
lblMail.Caption = ""
labelbybank.Caption = ""
maxrow = 0
addoredit = False
Unload frmEndPartTrans
       
       
End Sub
Public Function templost() As Boolean
    Dim check As Boolean
    Dim Row, Col As Integer
    Dim RRR, CCC As Integer
    Dim r, q, D As Double
    Dim mprevcol As Integer
    Dim mq As Currency, mr As Currency, mrot As Currency
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    If lastrow <= 0 Then
        templost = True
        Exit Function
    End If
    RRR = Grid1.Row
    
    CCC = Grid1.Col
    Grid1.Row = lastrow
    Grid1.Col = lastcol
    mprevcol = Grid1.Col
    Select Case Grid1.Col
            
       
            
            Case 1
                Grid1.text = tempmeb.text
                '/*************************
                'If RS.State = 1 Then
                '    RS.close
                'End If
                'RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                Set RS = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(Grid1.text) & "'")
                Row = Grid1.Row
                Col = Grid1.Col
                If Trim(Grid1.text) <> "" Then
                    If Not RS.BOF Then
                        'RS.MoveFirst
                        'RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF Then
                            tempmeb.Visible = True
                            tempmeb.SetFocus
                            RS.Close
                            templost = False
                            Exit Function
                        Else
                            Grid1.text = RS(0)
                            Grid1.Col = 2
                            Grid1.text = RS(1)
                         '   If Not edit Then
                                Grid1.Col = 3
                                If Trim(Grid1.text) = "" Then
                                    Grid1.text = 0
                                End If
                                q = Val(Grid1.text)
                                Grid1.Col = 5
                                If Trim(Grid1.text) = "" Then
                                    Grid1.text = Format(RS(3), "0.00")            'rs(3)
                                    r = RS(3)
                              
                                End If
                                '/******************
                                category = returnCategory(Trim(RS(2)))
                                If category = "C1" Then
                                  Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear)
                                ElseIf category = "C2" Then
                                  Set kk = con.Execute("select CATEGORY2 from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear)
                                ElseIf category = "C3" Then
                                  Set kk = con.Execute("select CATEGORY3 from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear)
                                End If
                                
                                Grid1.Col = 6
                                If Grid1.text = "" And addmode = True Then
                                    If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        kk.Close
                                        Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "' and " & stringyear)
                                        Grid1.Col = 4
                                        If kk.BOF Then
                                             GoTo abc
                                        End If
                                        Grid1.text = Format(kk(0), "0.00")
                                        Grid1.Col = 6
                                        Grid1.text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = RS(3)
                                    Else
abc:
                                        Grid1.Col = 4
                                        Grid1.text = Format(RS(4), "0.00")
                                        Grid1.Col = 6
                                        Grid1.text = Format(RS(4), "0.00")
                                        D = RS(4)
                                End If
                                
                                
                                
                                'New Code For Sub Group
                                    If IsEmpty(tempstr) Then tempstr = ""
                                    s_ = tempstr
                                    
                                    If rs1.State = 1 Then rs1.Close
                                    rs1.Open "select GROUPCODE_sub from books where bookcode='" & RS(0) & "'", con
                                    If rs1.EOF = False Then
                                    If (Not IsNull(rs1!GROUPCODE_sub) And Len(rs1!GROUPCODE_sub) > 0) Then
                                    
                                        D = ReturnDiscount("" & category, "" & s_, Trim(rs1(0)))
                                        If D > 0 Then
                                            Grid1.Col = 4
                                            Grid1.text = Format(D, "    0.00")
                                            Grid1.Col = 6
                                            Grid1.text = Format(D, "0.00")
                                            r = RS(3)
                                        End If
                                    End If
                                    End If
                                'End Code For Sub Group
                                
                                'Series Wise Discount
                                
                                
                                D = ReturnDiscountNew(RS(0), Trim(customercode.text), txtScId.text)
                                If D > 0 Then
                                    Grid1.Col = 4
                                    Grid1.text = Format(D, "    0.00")
                                    Grid1.Col = 6
                                    Grid1.text = Format(D, "0.00")
                                    r = RS(3)
                                End If
                                
                            
                                Grid1.Col = 7
                                Grid1.text = Format(Round(q * r, 2), "0.00")
                                Grid1.Col = 8
                                Grid1.text = Format(Round((q * r) * (D / 100), 2), "0.00")
                                
                              Else
                              
                                  If Grid1.text = "" And addmode = False Then
                                    If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        kk.Close
                                        Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "' and " & stringyear)
                                        Grid1.Col = 4
                                        If kk.BOF Then
                                             GoTo abc
                                        End If
                                        
                                        
                                        
                                        Grid1.text = Format(kk(0), "0.00")
                                        Grid1.Col = 6
                                        Grid1.text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = RS(3)
                                        
                                        
                                        
                                        
                                   'New Code For Sub Group
                                    If IsEmpty(tempstr) Then tempstr = ""
                                    s_ = tempstr
                                    
                                    If rs1.State = 1 Then rs1.Close
                                    rs1.Open "select GROUPCODE_sub from books where bookcode='" & RS(0) & "'", con
                                    If rs1.EOF = False Then
                                    If (Not IsNull(rs1!GROUPCODE_sub) And Len(rs1!GROUPCODE_sub) > 0) Then
                                        D = ReturnDiscount("" & category, "" & s_, Trim(rs1(0)))
                                        If D > 0 Then
                                            Grid1.Col = 4
                                            Grid1.text = Format(D, "    0.00")
                                            Grid1.Col = 6
                                            Grid1.text = Format(D, "0.00")
                                            r = RS(3)
                                        End If
                                    End If
                                    End If
                                  'End Code For Sub Group
                                        
                                   'series Wise Discount
                                   
                                   D = ReturnDiscountNew(RS(0), Trim(customercode.text), txtScId.text)
                                   If D > 0 Then
                                        Grid1.Col = 4
                                        Grid1.text = Format(D, "    0.00")
                                        Grid1.Col = 6
                                        Grid1.text = Format(D, "0.00")
                                        r = RS(3)
                                    End If
     
                                        
                                  End If
                                  End If
                              
                              End If
                          '  End If
                            Grid1.Col = Col
                            RS.Close
                        End If
                    End If
                End If
            Case 3, 5, 6
                If Grid1.Col <> 3 Then
                    Grid1.text = Format(Trim(tempmeb.text), "0.00")
                Else
                    Grid1.text = Format(Trim(tempmeb.text), "0")
                End If
                If Trim(Grid1.text) = "" Then
                    Grid1.text = 0
                End If
                Row = Grid1.Row
                Col = Grid1.Col
                
                Grid1.Col = 3
                q = Val(Trim(Grid1.text))
                Grid1.Col = 5
                r = Val(Trim(Grid1.text))
                Grid1.Col = 6
                D = Val(Trim(Grid1.text))
                Grid1.Col = 7
                Grid1.text = Format(Round(q * r, 2), "0.00")
                Grid1.Col = 8
                Grid1.text = Format(Round((q * r) * (D / 100), 2), "0.00")
                Grid1.Col = Col
            Case 4
                Grid1.text = tempmeb.text
                If Trim(Grid1.text) = "" Then
                    Grid1.text = 0
                End If
        End Select
        Row = Grid1.Row
        Col = Grid1.Col
        
        
        
        
        If (Col = 6 Or Col = 1) Then
        
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Round(Val(Trim(Grid1.text)), 2)
            Grid1.Col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.text)), 2)
        Next
        
        
        Me.tqu.Caption = ""
        For I = 1 To maxrow
            Grid1.Col = 3
            Grid1.Row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.text))
        Next
        
        End If
        
        
        invoicecalc
        
       
        
        
        Grid1.Row = RRR
        Grid1.Col = CCC
        templost = True
End Function

Private Sub bdated_LostFocus()
If Trim(bdated.text) <> Trim("__/__/____") Then
   If Not checkdate(Trim(bdated.text), bdated) Then
         bdated.SetFocus
    End If
End If
End Sub

Private Sub biltno_LostFocus()
biltno = UCase(biltno)
End Sub

Private Sub Bookcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If

End Sub
Function returnCategory(s As String) As String
    Dim s1 As New ADODB.Recordset
    If s1.State = 1 Then s1.Close
    
    s1.Open "select category from [groups] where groupcode='" & s & "' and " & stringyear & "", con
    If s1.EOF = False Then
       returnCategory = s1(0)
    End If
    
End Function
Private Sub Bookname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Dim mprevcol As Integer
        Dim mq As Currency, mr As Currency, mrot As Currency
        mprevcol = Grid1.Col
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        Select Case Grid1.Col
            Case 2
                Dim Row, Col As Integer
                Row = Grid1.Row
                Col = Grid1.Col
                If Trim(Bookname.text) = "" Then
                    Grid1.Col = 1
                    If Trim(Grid1.text) = "" Then
                        Grid1.text = Bookname.text
                           Bookname.SetFocus
  '********* vk
                          
                          
                          If Trim(Grid1.text) = "" And Row = 1 Then
                                 Grid1.Col = 2
                                 Grid1.text = ""
                                 If Trim(Grid1.text) = "" Then
                                           
                                          Grid1.Col = 1
                                          Bookname.SetFocus
                                          Grid1.SetFocus
                                       Exit Sub
                                 End If
                           End If
              '********
                           Commandother.SetFocus
                           Exit Sub
                    End If
                End If
                Grid1.Row = Row
                Grid1.Col = Col
                Grid1.text = Bookname.text
                '/*************************
                If RS.State = 1 Then
                    RS.Close
                End If
                Set RS = con.Execute("select * from books")
                 
                Row = Grid1.Row
                Col = Grid1.Col
                If Trim(Grid1.text) <> "" Then
                    If Not RS.BOF Then
                        RS.MoveFirst
                        RS.Find "bookname='" + Trim(Grid1.text) + "'"
                        If RS.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            RS.Close
                            Exit Sub
                        Else
                            
                            Grid1.Col = 1
                            Grid1.text = RS(0)
                            Grid1.Col = 2
                            Grid1.text = RS(1)
                        '   If Not edit Then
                                 Grid1.Col = 3
                            If Trim(Grid1.text) = "" Then
                                Grid1.text = 0
                            End If
                            q = Val(Grid1.text)
                            Grid1.Col = 5
                            Grid1.text = Format(RS(3), "0.00")
                            r = RS(3)
                            '/******************
                            
                            category = returnCategory(Trim(RS(2)))
                            If category = "C1" Then
                            Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear)
                            ElseIf category = "C2" Then
                            Set kk = con.Execute("select Category2 from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear)
                            ElseIf category = "C3" Then
                            Set kk = con.Execute("select Category3 from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear)
                            
                            End If
                            
                            Grid1.Col = 6
                            
                            If Trim(kk(0)) <> "" Then
                               tempstr = Trim(kk(0))
                               kk.Close
                               
                               Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "' and " & stringyear)
                               
                               
                               If kk.BOF Then
                                   GoTo abc
                               End If
                               
                               Grid1.Col = 4
                               Grid1.text = Format(kk(0), "0.00")
                               Grid1.Col = 6
                               Grid1.text = Format(kk(0), "0.00")
                               D = kk(0)
                            Else
abc:
                                 Grid1.Col = 4
                                 Grid1.text = Format(RS(4), "0.00")
                                    Grid1.Col = 6
                                    Grid1.text = Format(RS(4), "0.00")
                                    D = RS(4)
                                End If
                                
                                '===============================================
                                    'New Code For Sub Group
                                    If IsEmpty(tempstr) Then tempstr = ""
                                    s_ = tempstr
                                    
                                    If rs1.State = 1 Then rs1.Close
                                    rs1.Open "select GROUPCODE_sub from books where bookcode='" & RS(0) & "'", con
                                    If rs1.EOF = False Then
                                    If (Not IsNull(rs1!GROUPCODE_sub) And Len(rs1!GROUPCODE_sub) > 0) Then
                                        D = ReturnDiscount("" & category, "" & s_, Trim(rs1(0)))
                                        If D > 0 Then
                                            Grid1.Col = 4
                                            Grid1.text = Format(D, "    0.00")
                                            Grid1.Col = 6
                                            Grid1.text = Format(D, "0.00")
                                             r = RS(3)
                                        End If
                                    End If
                                    End If
                                  'End Code For Sub Group
                                 '===============================================
                                 
                                Grid1.Col = 7
                                Grid1.text = Round(q * r, 2)
                                Grid1.Col = 8
                                Grid1.text = Round((q * r) * (D / 100), 2)
                         '   End If
                            Grid1.Col = Col
                            RS.Close
                        End If
                    End If
                End If
        End Select
        Row = Grid1.Row
        Col = Grid1.Col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Round(Val(Trim(Grid1.text)), 2)
            Grid1.Col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.text)), 2)
        Next
        invoicecalc
        Grid1.Row = Row
        Grid1.Col = Col
        Select Case Grid1.Col
            Case 1
                Grid1.Col = 3
                Grid1.SetFocus
                Grid1_Click
            Case 2
                Grid1.Col = 3
                Grid1.SetFocus
                Grid1_Click
            Case 3, 4, 5
                Grid1.Col = Grid1.Col + 1
                Grid1.SetFocus
                Grid1_Click
            Case 6
                Grid1.Col = 1
                Grid1.Row = Grid1.Row + 1
                Grid1.SetFocus
                Grid1_Click
        End Select
    End If
End Sub
Private Sub Bookname_LostFocus()
    Bookname.Visible = False
End Sub
Private Sub bundles_LostFocus()
bundles = UCase(bundles)
End Sub

Private Sub Combosldistrictcode_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'        If Trim(Me.customercode.Text) <> "" Then
 '           Grid1.col = 1
 '           Grid1.row = 1
 '           Grid1_Click
 '       Else
 '           Me.textbox.SetFocus
 '           'Me.customercode.SetFocus
 '       End If
 '   End If
End Sub

Private Sub Check1_disZero_Click()

If Check1_disZero.value = 1 Then
  
   
    Dim totalamount_ As Double
    Dim totaldiscount_ As Double
    Dim qty
    totalamount_ = 0
    totaldiscount_ = 0
    Qty_ = 0
  
  
  For I = 1 To Grid1.rows - 1
     
    If (Grid1.TextMatrix(I, 1) <> "") Then
     Grid1.TextMatrix(I, 4) = 0
     Grid1.TextMatrix(I, 6) = 0
     Grid1.TextMatrix(I, 8) = 0
     
     
     totalamount_ = totalamount_ + Val(Grid1.TextMatrix(I, 7))
     totaldiscount_ = totaldiscount_ + Val(Grid1.TextMatrix(I, 8))

     
     
    End If
    
     
  Next
  
 
totalamount = totalamount_
totaldiscount = totaldiscount_

mga.Caption = Format(Round(totalamount_, 2), "0.00")
mgd.Caption = Format(Round(totaldiscount_, 2), "0.00")
mna.Caption = Format(Round((totalamount_ - totaldiscount_), 2), "0.00")
  
  

End If

End Sub

Private Sub Check1_spremarks_Click()
If Check1_spremarks.value = 1 Then
   txtPendingBooksRem.Enabled = True
   txtPendingBooksRem.SetFocus
   Label29.Enabled = True
Else
   txtPendingBooksRem.Enabled = False
   Label29.Enabled = False

End If
End Sub

Private Sub cmbAgentName_GotFocus()
'cmbAgentName.ListIndex = 0
End Sub
Private Sub cmbAgentName_KeyDown(KeyCode As Integer, Shift As Integer)
   
 If KeyCode = 13 Then
 
   If cmbAgentName.text = "" Then
    MsgBox "Select Reprasentative ...", vbCritical
    cmbAgentName.SetFocus
    Exit Sub
   End If
   
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
Sub fatchOrder()
      
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim dis
Dim party_ As String

party_ = ""

If PopUpValue1 = "" Then Exit Sub

If RS.State = 1 Then RS.Close
RS.Open "select partyname,invoiceno,ORDERBY,ORDERDATE,Transport,SUBLEDGER," & _
"repname,ScName,ScId,Shipto_CityId,Shipto,Godown,address1,address2,party_dist" & _
",party_state,Shipto_Add1,Shipto_Add2,shipto_dist,Shipto_States,Frt_Yes  from ordera where sale_sp='sale' and invoiceno=" & PopUpValue1 & " and " & stringyear, con
If RS.EOF = False Then

      lblPartyfrt.Caption = RS!Frt_Yes
      I_OB = RS!orderby
      txtOrderNo = RS!invoiceNo
      I_DTOB.text = RS!ORDERDATE
      textbox = RS!partyname
      customercode = RS!partyname
      cmbtransportname.text = RS!transport
      cmbAgentName.text = RS!RepName & ""
      txtschool.text = RS!scname & ""
      txtScId.text = RS!scid & ""
      
      If party_ = "" Then
         party_ = RS!Shipto_Add1
      End If
      
      If party_ = "" Then
         party_ = RS!Shipto_Add2
      Else
         party_ = party_ & ", " & RS!Shipto_Add2
      End If
      
      
      If party_ = "" Then
         party_ = RS!shipto_dist
      Else
         party_ = party_ & ", " & RS!shipto_dist
      End If
      
      If party_ = "" Then
         party_ = RS!Shipto_States
      Else
         party_ = party_ & ", " & RS!Shipto_States
      End If
      
      txtShip = RS!Shipto & ", " & party_
      
      'lblBookSId = RS!Shipto_CityId & ""
      txtMark.text = RS!Godown & ""
      
Else
Exit Sub
      
      
End If

Dim kk1 As Integer
Dim sqty As Integer
Dim oqty
Dim b1 As Boolean
b1 = False
sqty = 0
oqty = 0
kk1 = 1

If txtOrderNo = "" Then Exit Sub

If rs3.State = 1 Then rs3.Close
rs3.Open "select top 1 OrderNo from PackingQry where OrderNo=" & txtOrderNo & "", con
If rs3.EOF = True Then
   b1 = True
Else
   b1 = False
End If



'---------SaleBookList
If RS.State = 1 Then RS.Close
RS.Open "select bookCode,bookname,rate,quantity,unit,amount,printorder,discount,rateMain from OrderBookList where sale_sp='sale' and invoiceno=" & PopUpValue1 & " order by printorder", con
For I = 1 To RS.RecordCount
      
    ''============================================
    sqty = 0
    oqty = 0
    str1 = "select top 100 BookCode,BookName,BillNo,sum(Qty) from PackingQry where (OrderNo=" & txtOrderNo & " and bookcode='" & RS(0) & "') group by BookCode,BookName,BillNo"
    
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open str1, con, adOpenForwardOnly, adLockReadOnly
    If rs1.EOF = False Then
    
    While rs1.EOF = False
     oqty = oqty + rs1(3)
     'I_NO = rs1(2)
     rs1.MoveNext
    Wend
    
     If rs2.State = 1 Then rs2.Close
     rs2.Open "select sum(QUANTITY) from invoiceBQry where OrderNo=" & txtOrderNo & " and bookcode='" & RS!Bookcode & "'", con
     If Not IsNull(rs2(0)) Then
        sqty = rs2(0)
     End If
    End If
    
    ''===========================================
    If (oqty - sqty) > 0 Then
    If b1 = False Then
      Grid1.TextMatrix(kk1, 1) = RS(0)
      Grid1.TextMatrix(kk1, 2) = RS(1)
      Grid1.TextMatrix(kk1, 3) = (oqty - sqty)
      
      dis = ReturnDiscountNew(RS(0), textbox, txtScId.text)
      
      '''Grid1.TextMatrix(kk1, 4) = RS!discount
      Grid1.TextMatrix(kk1, 4) = dis
      
      Grid1.TextMatrix(kk1, 5) = RS!rateMain
      
      '''Grid1.TextMatrix(kk1, 6) = RS!discount
      Grid1.TextMatrix(kk1, 6) = dis
      Grid1.TextMatrix(kk1, 7) = (Grid1.TextMatrix(kk1, 3) * RS!rate)
      '''Grid1.TextMatrix(kk1, 8) = Format(Round(Grid1.TextMatrix(kk1, 7) * (RS!discount / 100), 2), "0.00")
      Grid1.TextMatrix(kk1, 8) = Format(Round(Grid1.TextMatrix(kk1, 7) * (dis / 100), 2), "0.00")
      kk1 = kk1 + 1
    End If
    
   End If
    
    
    
  'If Packing Is not Exist======================================================================
  '=============================================================================================

If b1 = True Then

dinesh:

     If rs2.State = 1 Then rs2.Close
     rs2.Open "select sum(QUANTITY) from invoiceBQry where OrderNo=" & txtOrderNo & " and bookcode='" & RS!Bookcode & "'", con
     If Not IsNull(rs2(0)) Then
        sqty = rs2(0)
     Else
        sqty = 0
     End If
    
    oqty = RS!QUANTITY
    
    ''===========================================
    If (oqty - sqty) > 0 Then

      Grid1.TextMatrix(kk1, 1) = RS(0)
      Grid1.TextMatrix(kk1, 2) = RS(1)
      Grid1.TextMatrix(kk1, 3) = (oqty - sqty)
      dis = ReturnDiscountNew(RS(0), textbox, txtScId.text)
      Grid1.TextMatrix(kk1, 4) = dis           'RS!discount
      Grid1.TextMatrix(kk1, 5) = RS!rateMain   'RS!rate
      Grid1.TextMatrix(kk1, 6) = dis           'RS!discount
      Grid1.TextMatrix(kk1, 7) = (Grid1.TextMatrix(kk1, 3) * RS!rate)
      '''Grid1.TextMatrix(kk1, 8) = Format(Round(Grid1.TextMatrix(kk1, 7) * (RS!discount / 100), 2), "0.00")
      Grid1.TextMatrix(kk1, 8) = Format(Round(Grid1.TextMatrix(kk1, 7) * (dis / 100), 2), "0.00")
      kk1 = kk1 + 1
   End If

 End If
    
    
    
    
    
 
  
  '=============================================================================================
  RS.MoveNext
Next
     
     
     
     
     
     
     

Dim totalamount_ As Double
Dim totaldiscount_ As Double
Dim qty
totalamount_ = 0
totaldiscount_ = 0
Qty_ = 0

For I = 1 To Grid1.rows - 1
If Grid1.TextMatrix(I, 1) <> "" Then
   totalamount_ = totalamount_ + Val(Grid1.TextMatrix(I, 7))
   totaldiscount_ = totaldiscount_ + Val(Grid1.TextMatrix(I, 8))
   qty = qty + Val(Grid1.TextMatrix(I, 3))
End If
Next

mga.Caption = Format(Round(totalamount_, 2), "0.00")
mgd.Caption = Format(Round(totaldiscount_, 2), "0.00")
mna.Caption = Format(Round((totalamount_ - totaldiscount_), 2), "0.00")
Me.tqu.Caption = qty
    
PopUpValue1 = ""
      
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
Set RS = con.Execute("exec searchList 'I'")

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

Private Sub cmdMaster_Click()
party_name = textbox.text
frmSubledger.Show
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPendingClear_Click()
If txtOrderNo = "" Then Exit Sub
If MsgBox("Are you Sure ", vbQuestion + vbYesNo) = vbYes Then
  con.Execute "update ORDERA set BendingBill='y' where INVOICENO=" & txtOrderNo & ""
End If
End Sub
Private Sub cmdPendingOrder_Click()
'aaa = Format(txtcreatedDT.text, "MM/dd/yyyy HH:M:SS")
          


searchType = "inv"
popuplist10 "select InvoiceNo as OrderNo,InvoiceDate as OrderDate,Subledger as Party,NetAmount from orderA where BendingBill='n' and " & stringyear & "  order by InvoiceNo", con
End Sub
Private Sub cmdPendingOrder_GotFocus()
If PopUpValue1 <> "" Then
   
   'txtOrderNo = PopUpValue1
   
   invoiceabandon
   fatchOrder
   On Error Resume Next
    Dim ctl As Control
    For Each ctl In Me.Controls
    If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
        If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) Then
        End If
        ctl.Enabled = True
    End If
    Next

   
End If

   PopUpValue1 = ""
End Sub

Private Sub cmdref_Click()
Screen.MousePointer = vbHourglass
popupvalue5 = txtOrderNo.text
frmINVOrder.Show
popupvalue5 = ""
Screen.MousePointer = vbDefault
End Sub

Private Sub CommandDirectPrint_Click()
s1 = "101"
'PrintOption.Show



con.Execute "update a set a.shipContactNo=b.ContactNo   FROM INVOICEA as a inner join ORDERA b on (a.OrderNo = b.invoiceno) where (b.Sale_sp ='sale' and a.INVOICENO = " & I_NO & ")"

If Check1_dos.value = 1 Then
   printButton = "2"
   printinvoice
   PrintOption.Show
Else
   printButton = "1"
   PrintOption.Show
End If

''If option_withHeader.value = 1 Then
''   Screen.MousePointer = vbDefault
''   popupvalue5 = invoice.I_NO
''   popupvalue4 = "Bill.rpt"
''   frmSendMail.Show 1
''Else
''   PrintOption.Show
''End If



End Sub

Private Sub Commandprintnh_Click()
    printch = "INVOICEA"
    ino = I_NO
    printch1 = "INVOICENO"


Printheader = False
printinvoice
End Sub

Private Sub Commandabandon_Click()
invoiceabandon

Me.Commandall.Enabled = False
Me.Commandother.Enabled = False
 
 
mnuMenu_ = "menusalesinvoice"
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

End Sub
Private Sub Commandadd_Click()
    On Error Resume Next
    invoiceabandon
    Dim RS As ADODB.Recordset
    addoredit = True
    addmode = True
    Set RS = New ADODB.Recordset
    Dim TEMPNUM As Integer
    
    If Edit = False Then
       Me.I_NO.text = con.Execute("Select max(invoiceno) from invoicea where " & stringyear)(0) + 1
    End If
    
    Dim ctl As Control
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = True
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(Me.I_NO.Name)) Then
           'ctl.Enabled = False
        End If
    Next
    
    txtTODNO.Enabled = False
    txtTODDate.Enabled = False
    Check1_trans.Enabled = False
    
    Me.Edit = False
    Picture5.Enabled = True
    Commandother.Enabled = True
    Commandadd.Enabled = False
    Commanddelete.Enabled = False
    Commandedit.Enabled = False
    CommandPrint.Enabled = False
     Commandprintnh.Enabled = False
    Commandall.Enabled = True
    Commandsave.Enabled = True
    Commandsearch.Enabled = False
    Grid1.Enabled = True
    Me.customercode.Enabled = True
    Check1_school.Enabled = True
    
    addoredit = True
    I_NO.SetFocus
End Sub
Private Sub Commandall_Click()
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Dim myvalue As String

If Trim(Me.customercode.text) = "" Then
    MsgBox "Please Fill the customer detail "
    Exit Sub
End If

myvalue = InputBox("Please enter the quantity ", "Enter the quantity: ", "1")
    
If Len(myvalue) > 0 And Val(myvalue) > 0 Then
    
    Grid1.rows = 1
    Grid1.rows = 2
    Grid1.Col = 1
    Grid1.Row = 1
    If RS.State = 1 Then
        RS.Close
    End If
    RS.Open "select * from books order by BOOKCODE", con, adOpenDynamic, adLockReadOnly, adCmdText
    Row = Grid1.Row
    Col = Grid1.Col
    If Not RS.BOF Then
        RS.MoveFirst
        Do While Not RS.EOF
            Grid1.Col = 1
            Grid1.text = RS(0)
            Grid1.Col = 2
            Grid1.text = RS(1)
            Grid1.Col = 3
            If Trim(Grid1.text) = "" Then
                Grid1.text = Val(myvalue)
            End If
            q = Val(Grid1.text)
            Grid1.Col = 5
            Grid1.text = Format(RS(3), "0.00")
            r = RS(3)
            
            '/******************
            'Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
            category = returnCategory(Trim(RS(2)))
            If category = "C1" Then
               Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear)
            ElseIf category = "C2" Then
               Set kk = con.Execute("select Category2 from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear)
            ElseIf category = "C3" Then
               Set kk = con.Execute("select Category3 from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear)
            End If
             
            Grid1.Col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.Close
                Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "' and " & stringyear)
                Grid1.Col = 4
                If kk.BOF Then
                    GoTo abc
                End If
                Grid1.text = Format(kk(0), "0.00")
                Grid1.Col = 6
                Grid1.text = Format(kk(0), "0.00")
                D = kk(0)
            Else
abc:
                Grid1.Col = 4
                Grid1.text = Format(RS(4), "0.00")
                Grid1.Col = 6
                Grid1.text = Format(RS(4), "0.00")
                D = RS(4)
            End If
            Grid1.Col = 7
            Grid1.text = Format(Round(q * r, 2), "0.00")
            Grid1.Col = 8
            Grid1.text = Format(Round((q * r) * (D / 100), 2), "0.00")
            If Not RS.EOF Then
                Grid1.rows = Grid1.rows + 1
                Grid1.Row = Grid1.Row + 1
                RS.MoveNext
            End If
        Loop
    
    '/**fghfghgh
    
    End If
    RS.Close
    totalamount = 0
    totaldiscount = 0
    Me.tqu.Caption = ""
    For I = 1 To Grid1.rows - 1
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Round(Val(Trim(Grid1.text)), 2)
            Grid1.Col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.text)), 2)
            Grid1.Col = 3
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.text))
     Next
     maxrow = Grid1.rows - 1
Else
Exit Sub
End If

invoicecalc

End Sub
Private Sub Commanddelete_Click()
On Error GoTo Del

'======================================

Dim rs_h As New ADODB.Recordset
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset


If rs1.State = 1 Then rs1.Close
rs1.Open "select top 1 * from invoicea where invoiceno=" & I_NO.text & " and " & stringyear, con, adOpenStatic, adLockReadOnly
If rs1.EOF = False Then
    If rs1!bAuthorized = True Then
        MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
        Exit Sub
    End If
   
End If


'=======================================

createLog UserName, I_NO, "invoice ", " Dalete : " & mna.Caption, Date


If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else

    
    
    If (AuditTrail = "y") Then
    
    If (txtchecked.text = "y") Then
    
        actionType_ = "Delete"
        vtype1_ = "I"
        vtypeNew = "I"
        vdate_ = Trim(i_dt.text)
        vno_ = Trim(I_NO.text)
        
        frmAuditTrailLog_Rem.Show 1
        
     End If
    
    End If



    con.Execute ("delete  from invoicea where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
    con.Execute ("delete  from INVOICEB where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
    con.Execute ("delete  from invoicec where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
    con.Execute ("delete  from invoiceb_Free where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
    
    invoiceabandon
    
    
End If


Commanddelete.Enabled = False
Commandadd.SetFocus

Exit Sub
Del:
MsgBox "" & err.Description

End Sub

Private Sub Commandedit_Click()
    
    Commandadd.Enabled = False
    Me.Commandedit.Enabled = False
    Picture5.Enabled = True
    Commandadd.Enabled = False
    Commandedit.Enabled = False
    Commandall.Enabled = True
    Commandsave.Enabled = False
    Commanddelete.Enabled = False
    Commandsearch.Enabled = False
    CommandPrint.Enabled = False
    Commandprintnh.Enabled = False
    Check1_disZero.Enabled = True
    
    Grid1.Enabled = True
    Commandall.Enabled = False
    Me.customercode.Enabled = True
    Edit = True
    addoredit = False
    
    I_NO_LostFocus
    i_dt.Enabled = True
    i_dt.SetFocus
    
    'CON.Execute ("delete  from invoicectmp WHERE " & stringyear & " and username='" & username & "' and INVOICENO = " + Trim(I_NO.Text))
    con.Execute ("delete  from invoicectmp WHERE " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
    DoEvents
    con.Execute ("insert into invoicectmp([INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],[saleType],username) " & _
    "  select [INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],[saleType],username " & _
    " from invoicec where  " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
    DoEvents
    
    Dim kx As Integer
    kx = 0
    
    addoredit = False
    
    HIT
    
    
    panel.Enabled = True
    Me.Enabled = True
    
    
    Dim ctl As Control
    For Each ctl In Me.Controls
    
    If (TypeOf ctl Is Label Or TypeOf ctl Is textbox Or TypeOf ctl Is MaskEdBox Or TypeOf ctl Is ComboBox) Then
        ctl.Enabled = True
    End If
    
    Next
    
    
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    Commandsave.Enabled = False
    'Commanddelete.Enabled = True
    Check1_school.Enabled = True
    
End Sub
Private Sub Commandother_Click()

mnuMenu_ = "menusalesinvoice"
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

Commandsave.Enabled = True
searchForm = "invoice"
frmEndPartTrans.Show
frmEndPartTrans.Refresh
DoEvents
DoEvents
DoEvents
DoEvents
 
End Sub
Private Sub CommandPrint_Click()
  
  On Error GoTo pp
  
  If Check1_trans.value = 1 Then
     con.Execute "exec salepending_orderwise '" & txtOrderNo & "'"
     con.Execute "update a set a.shipContactNo=b.ContactNo   FROM INVOICEA as a inner join ORDERA b on (a.OrderNo = b.invoiceno) where (b.Sale_sp ='sale' and a.INVOICENO = " & I_NO & ")"
  End If
  
  
  printch = "INVOICEA"
  ino = I_NO
  printch1 = "INVOICENO"
  s1 = "1"
  Printheader = True
  
  If Check1_dos.value = 1 Then
     printButton = "2"
     printinvoice
  Else
     printButton = "1"
  End If
  
  If option_withHeader.value = 1 Then
     Screen.MousePointer = vbDefault
     popupvalue5 = invoice.I_NO
     popupvalue4 = "Bill.rpt"
     frmSendMail.Show 1
  Else
     
     If Check1_withheader.value = 1 Then
        PopUpValue6 = "withheader"
     Else
        PopUpValue6 = ""
     End If
  
     PrintOption.Show
  End If
  
  
Exit Sub
pp:
MsgBox "" & err.Description
   
End Sub
Private Sub CommandReturn_Click()
   
'''   Dim RS As New ADODB.Recordset
'''   RS.Open "tempINV", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
'''   If RS.BOF Then
'''       RS.AddNew
'''   End If
'''   RS!In = CON.Execute("Select max(invoiceno) from INVOICEA")(0)
'''   RS.Update
'''   RS.Close
   Unload Me
'''   addoredit = False
'''   'MainMenu.Toolbar1.Visible = True
End Sub
Function checkOrderQty() As Boolean

   Dim oqty
   Dim BQty
   Dim qty
   
   
   If Edit = True Then
    'con.Execute ("delete  from invoicea where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
    'con.Execute ("delete  from INVOICEB where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
    'con.Execute ("delete  from invoicec where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
    'con.Execute ("delete  from invoiceb_Free where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
   End If
   
   '================================================
   For d1 = 1 To Grid1.rows - 1

   oqty = 0
   BQty = 0
   qty = 0


   If Grid1.TextMatrix(d1, 1) <> "" Then
   con.Execute "exec sale_pendingSp '" & txtOrderNo.text & "' "
   Set rs1 = con.Execute("exec checkOrderQty'" & Trim(txtOrderNo.text) & "','" & Grid1.TextMatrix(d1, 1) & "'")
    If rs1.EOF = False Then
       oqty = rs1!oqty
       BQty = rs1!BQty
    End If

    If Val(Grid1.TextMatrix(d1, 3)) > 0 Then
       BQty = BQty + Val(Grid1.TextMatrix(d1, 3))
    End If

    If oqty < BQty Then
       checkOrderQty = True
       'MsgBox "Qty. Exceed Related Order .... ", vbCritical
       Exit Function
    End If

   End If

   Next

   '================================================


End Function
Sub updateRandomId(inv As String)
s10 = ""

Set kk1 = New ADODB.Recordset
kk1.Open "select invoiceno,RandomNo FROM INVOICEA where invoiceno=" & inv & "", con
If kk1.EOF = False Then

If Len(kk1!RandomNo) >= 5 Then
   s10 = Mid(value_, 1, 2)
Else
   s10 = Mid(value_, 1, 3)
End If

s10 = "IN" & s10 & kk1!RandomNo

con.Execute "update INVOICEA set Randomid= '" & s10 & "'  where invoiceno=" & kk1!invoiceNo & ""

End If

End Sub
Private Sub Commandsave_Click()
    
On Error GoTo save_
    
    
    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    
    Dim Checked_YesNo  As Integer
   
    If (txtchecked.text = "y") Then
         Checked_YesNo = 1
    Else
         Checked_YesNo = 0
    End If
  
  
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select top 100 * from invoicea where invoiceno=" & I_NO.text & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.Close
       rs_h.Open "select top 100 * from invoicea where invoiceno=" & I_NO.text & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
           If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
          End If
       
    End If
    
    createLog UserName, I_NO, "invoice ", " Save : " & mna.Caption, Date
    
   
    
    
    If Edit = False Then
    din1 = checkPacking(I_NO, "inv")
    
    If (din1 > Val(tqu) Or din1 < Val(tqu)) Then
      If MsgBox("Packing Quantity is " & din1 & vbCrLf & " It is differ Quantity to this bill, Are Sure to Continue.. ", vbQuestion + vbYesNo) = vbNo Then
         Exit Sub
      End If
    End If
    
    End If
    '--------------------------------
        
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
     
    Dim SAVED As Boolean
    Dim LAMOUNT As Double
    If Edit = False And addmode = False Then
      Me.Commandsave.Enabled = False
      Exit Sub
    End If
    
     
    
    If MsgBox("Do you want to save it now ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    Grid1.Row = 1
    Grid1.Col = 1
    
    If Trim(Grid1.text) = "" Then
       MsgBox "Please Enter item.... "
       Exit Sub
    End If
    
    If Trim(cmbAgentName.text) = "" Then
       MsgBox "Please Select Reprasentative.... "
       cmbAgentName.SetFocus
       Exit Sub
    End If
    
    
    
    SAVED = False
    
    
    If Edit = False Then
       If check_Duplikate("invoicea", I_NO.text) = True Then
           If con.Execute("Select max(invoiceno) from invoicea")(0) >= Val(Trim(Me.I_NO.text)) Then
                Me.I_NO.text = Str(Val(Trim(Me.I_NO.text)) + 1)
           End If
         'Exit Sub
       End If
       
       
    End If
    
    
    If Trim(I_NO.text) <> "" And Trim(i_dt.text) <> "" And Trim(customercode.text) <> "" Then
            
            If Edit Then
                
                con.Execute ("delete  from invoicea where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
                con.Execute ("delete  from INVOICEB where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
                con.Execute ("delete  from invoicec where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
                con.Execute ("delete  from invoiceb_Free where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
                
                
           
            
            End If
            
            
            '================================
            If Val(txtOrderNo) > 0 Then
            If checkOrderQty = True Then
               MsgBox "Qty. Exceed Related Order .... ", vbCritical
               Exit Sub
            End If
            End If
            '================================
            
            
            If RS.State = 1 Then
                RS.Close
            End If
            LAMOUNT = 0
            
    'Code for Order Mnm
    
    If Edit = False Then
          'If (I_OB <> "" And txtMark <> "") Then
          'Party_Remove_FromOrder Trim(Me.customercode.Text), txtMark, Trim(I_OB)
          'End If
    End If
    
            
    RS.Open "select * from invoicea where " & stringyear & " and invoiceno <=0", con, adOpenDynamic, adLockOptimistic
    If Not Edit Then
again:
    End If
            
            
            
            RS.AddNew
            
            
            If (AuditTrail = "y") Then
            RS!Checked_YesNo = Checked_YesNo
            End If
            
            RS!PendingRemarks = txtPendingBooksRem.text
            RS!app_add = lblApp(4).Caption
            RS!appno = txtappno.text
            
            If txtTODNO.text <> "" Then
               RS!todid = txtTODNO.text
               RS!toddate = txtTODDate.text
            Else
               RS!todid = Null
               RS!toddate = Null
            
            End If
            
            
            If Check1_school.value = 1 Then
            RS!Shipto_Scholl = 1
            Else
            RS!Shipto_Scholl = 0
            End If
            
            RS!Placeofsupply = Trim(txtPlaceSup)
            RS!remarks = Trim(txtRem)
            RS!NsChallanNo = Trim(txtNSCHNo)
            
            
            If Len(txtShip) > 0 Then
            If Len(lblBookSId) > 0 Then
              If InStr(Trim(txtShip), ",") > 0 Then
               RS!Shipto = Mid(Trim(txtShip), 1, InStr(Trim(txtShip), ",") - 1)
              Else
               RS!Shipto = Trim(txtShip)
              End If
               
            Else
               RS!Shipto = Trim(txtShip)
            End If
            End If
            
            
            If Trim(Me.txtRandomDT) <> Trim("__/__/____") Then
                RS!SMSDate = Trim(Me.txtRandomDT.text)
            End If
            RS!randomId = Trim(txtRandomId.text)
            RS!mobile = Trim(txtRandomMob.text)
            
 
            
            RS!Shipto_CityId = Trim(lblBookSId.Caption)
            RS!scname = Trim(txtschool)
            RS!scid = Trim(txtScId.text)
            RS!invoiceNo = Val(Me.I_NO.text)
            RS!invoiceDate = Me.i_dt.text
            RS!Genledger = Trim(Me.Genledger.text)
            RS!subledger = Trim(Me.customercode.text)
            RS!agentname = Trim(Me.cmbAgentName.text)
            RS!transportname = Trim(Me.cmbtransportname.text)
            RS!orderby = Trim(Me.I_OB.text)
            RS!orderNo = Val(txtOrderNo.text)
            
            If Trim(Me.I_DTOB) <> Trim("__/__/____") Then
            '    rs!ORDERDATE = Date
            'Else
                RS!ORDERDATE = Trim(Me.I_DTOB.text)
            End If
            RS!marka = Trim(Me.marka.text)
            RS!Godown = IIf(txtMark.text = "", "n", txtMark.text)
            RS!bundles = Trim(Me.bundles)
            RS!through = Trim(Me.through.text)
            RS!through1 = Trim(Me.through1.text)
            If Trim(Me.through1.text) = "" Then
                RS!through1 = " "
            End If
            RS!station = Trim(Me.station.text)
            RS!biltyno = Trim(Me.biltno.text)
            If Trim(Me.bdated) <> Trim("__/__/____") Then
                RS!BILTYDATE = Me.bdated & ""
           End If
            
            RS!freight = Trim(Me.freight)
            RS!weight = Trim(Me.weight)
            RS!netamount = Round(Val(Trim(Me.mna.Caption)), 2)
            RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
            RS!txt1 = Trim(frmEndPartTrans.T1TEXT.text)
            RS!txt1a = Val(Trim(frmEndPartTrans.T1.text))
            RS!txt2 = Trim(frmEndPartTrans.T2TEXT.text)
            RS!txt2a = Val(Trim(frmEndPartTrans.T2.text))
            RS!baa = Val(Trim(frmEndPartTrans.T3TEXT.text))
            RS!baa = Val(Trim(labelbybank.Caption))
            If addmode = True Then
                If Val(Trim(frmEndPartTrans.T3TEXT.text)) <> 0 Then
                      RS!advicestatus = "Pending"
                      Me.txtadst.text = "Pending"
                End If
            Else
                RS!advicestatus = Me.txtadst.text & ""
            End If
            Dim trs As New ADODB.Recordset
            trs.Open " SELECT DISTCODE FROM SLEDGER  WHERE SUBLEDGER='" & customercode.text & "' and " & stringyear, con, adOpenStatic, adLockOptimistic, adCmdText
            If Not trs.BOF Then
                RS!District = Trim(trs!distcode)
            Else
                RS!District = ""
            End If
err1:
           If Not Edit Then
                If con.Execute("Select max(invoiceno) from invoicea where " & stringyear & "")(0) >= Val(Trim(Me.I_NO.text)) Then
                    On Error GoTo err1
                End If
            End If
            RS!fyear = session
            RS!setupid = setupid
            RS!Amtwords = txtAmtwords.text
            
''            If lblBookSId.Caption <> "" Then
''                If kk.State = 1 Then kk.close
''                kk.Open "select add1,add2,City,District,State from QryBookSeller where BookSelerID='" & lblBookSId.Caption & "'", con
''                If kk.EOF = False Then
''                   RS!Shipto_Add1 = Trim(kk!add1)
''                   RS!Shipto_Add2 = Trim(kk!add2)
''                   RS!Shipto_City = Trim(kk!city)
''                   RS!Shipto_district = Trim(kk!District)
''                   RS!Shipto_States = Trim(kk!State)
''                   RS!Shipto_CityId = lblBookSId.Caption
''                End If
''            End If
            
            
            RS.update
            
            
            
            On Error GoTo 0
            RS.Close
            RS.Open "select * from INVOICEB where " & stringyear & " and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
            Dim I As Integer
            RRRR = Grid1.Row
            CCCC = Grid1.Col
            For I = 1 To maxrow
                Grid1.Row = I
                Grid1.Col = 1
                If Trim(Grid1.text) <> "" Then
                    Grid1.Col = 3
                    If Val(Trim(Grid1.text)) > 0 Then
                       Grid1.Col = 5
                       If Val(Trim(Grid1.text)) > 0 Then
                         RS.AddNew
                         Grid1.Col = 1
                         RS!invoiceNo = Val(Me.I_NO.text)
                         RS!invoiceDate = Me.i_dt.text
                         RS!Genledger = Trim(Me.Genledger.text)
                         RS!subledger = Trim(Me.customercode.text)
                         RS!Bookcode = Trim(Grid1.text)
                         Grid1.Col = 3
                         RS!QUANTITY = Trim(Grid1.text)
                         Grid1.Col = 5
                         RS!rate = Trim(Grid1.text)
                         Grid1.Col = 7
                         RS!amount = Trim(Grid1.text)
                         LAMOUNT = Val(Trim(Grid1.text))
                         Grid1.Col = 4
                         RS!PRINTORDER = Trim(Grid1.text)
                         Grid1.Col = 6
                         RS!discount = Trim(Grid1.text)
                         Grid1.Col = 8
                         RS!netamount = LAMOUNT - Trim(Grid1.text)
                         LAMOUNT = 0
                         RS!agentname = Trim(Me.cmbAgentName.text)
                         
                         RS!fyear = session
                         RS!setupid = setupid
                        
                        If kk.State = 1 Then kk.Close
                        kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,Apply from KitQry where kitcode='" & RS!Bookcode & "'", con
                        bkdesc = ""
                        While kk.EOF = False
                        
                        If kk!Apply = "y" Then
                           con.Execute "insert into INVOICEB_Free(INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
                           " values('" & Val(Me.I_NO.text) & "','" & Format(Me.i_dt.text, "MM/dd/yyyy") & "','" & Trim(Me.Genledger.text) & "','" & Trim(Me.customercode.text) & "','" & kk!Bookcode & "','" & (kk!qty * RS!QUANTITY) & "','" & kk!rate & "','" & Trim(Me.cmbAgentName.text) & "','" & setupid & "','" & session & "','" & txtMark & "')"
                        End If
                        
                        If bkdesc = "" Then
                          If kk!NoPrintDesc = False Then
                            bkdesc = kk!Bookname
                          End If
                          
                         Else
                           If kk!NoPrintDesc = False Then
                            bkdesc = bkdesc & "," & kk!Bookname
                           End If
                         End If
                         kk.MoveNext
                        Wend
                        
                        
                        If bkdesc <> "" Then
                        RS!BookDesc = "(" & bkdesc & ")"
                        End If
                        
                        RS!app_add = lblApp(4).Caption
                        RS!appno = txtappno.text
                        RS.update

                         
                       End If
                    End If
                End If
            Next
            RS.Close
            Grid1.TopRow = 1
            Grid1.Row = 1
            Grid1.Col = 1
            RS.Open "select * from invoicec where " & stringyear & " and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
            '/******
                'Dim I, x As Integer
                Dim temprs As ADODB.Recordset
                Set temprs = New ADODB.Recordset
                   
                With frmEndPartTrans
                
                
                For I = 1 To .vs.rows - 1
                   
                   If Trim(.vs.TextMatrix(I, 0)) <> "" Then
                   
                        
                        RS.AddNew
                        RS!fyear = session
                        RS!setupid = setupid
                        RS!UserName = UserName
                        
 
                        RS!invoiceNo = Val(Me.I_NO.text)
                        RS!invoiceDate = Me.i_dt.text
                        RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
                        RS!text = Trim(.vs.TextMatrix(I, 0))
                        If temprs.State = 1 Then
                            temprs.Close
                        End If
                        
                        
                        If Edit Then
                        ''temprs.Open "select * from invoicectmp WHERE username='" & username & "' and INVOICENO=" & invoice.I_NO & " and " & stringyear & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
                        temprs.Open "select * from invoicectmp WHERE INVOICENO=" & invoice.I_NO & " and " & stringyear & "", con, adOpenDynamic, adLockReadOnly, adCmdText
                        If .vs.TextMatrix(I, 0) <> "" Then
                                temprs.Find "TEXT='" + Trim(.vs.TextMatrix(I, 0)) + "'"
                                If temprs.EOF = False Then
                                RS!Genledger = Trim(temprs!Genledger)
                                RS!subledger = Trim(temprs!subledger)
                                RS!DebitorCredit = Trim(temprs!DebitorCredit)
                                RS!RYN = temprs!RYN & ""
                                End If
                                
                        End If
                        
                        temprs.Close
                        
                        
                        
                Else
                        
                        temprs.Open "select * from INVOICEEND where  type='invoice' and " & stringyear & " order by printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
                        If .vs.TextMatrix(I, 0) <> "" Then
                                temprs.Find "TEXT='" + Trim(.vs.TextMatrix(I, 0)) + "'"
                                RS!Genledger = Trim(temprs!Genledger)
                                RS!subledger = Trim(temprs!subledger)
                                RS!DebitorCredit = Trim(temprs!DebitorCredit)
                                RS!RYN = temprs!RYN & ""
                        End If
                        temprs.Close
                        End If
                        
                        RS!rate = Val(Trim(.vs.TextMatrix(I, 1)))
                        If Val(Trim(.vs.TextMatrix(I, 1))) > 0 Then
                            RS!amount = Round((Me.totalamount - Me.totaldiscount), 2) * Round((Val(Trim(.vs.TextMatrix(I, 1))) / 100), 2)
                        Else
                            RS!amount = Val(Trim(.vs.TextMatrix(I, 2)))
                        End If
                    RS.update
                    End If
            Next
            
            End With
            
            RS.Close
            
                
            If txtcreatedDT.text <> "" Then
               'con.Execute "update invoicea set CreatedDt = '" & Format(txtcreatedDT.text, "MM/dd/yyyy") & "' where invoiceno=" & I_NO.text & ""
               con.Execute "update invoicea set CreatedDt = '" & Format(txtcreatedDT.text, "MM/dd/yyyy HH:M:SS") & "' where invoiceno=" & I_NO.text & ""

               
               'YYYY-MM-DD HH:M:SS
            
            End If
            
            
            SAVED = True
        
        End If
             
            s11 = ""
            ss11 = ""
            
            s11 = InStr(1, Me.station.text, " ")
            If s11 <> 0 Then
            ss11 = Trim(Mid(Me.station.text, 1, s11))
            Else
            ss11 = Me.station.text
            End If
            PopUpValue1 = ss11
   
             
             UpdateDisPatchReg I_NO, i_dt, Me.customercode, PopUpValue1, Trim(Me.bundles), Trim(Me.cmbtransportname.text), Trim(Me.marka.text), Trim(Me.biltno.text), Me.bdated, Trim(Me.freight), "DispatchRegister"
             PopUpValue1 = ""
         ' End If

        
        If SAVED Then
            
            updateRandomId I_NO.text
            
            MsgBox "Record Saved"
            
            
            '========================================
        If txtOrderNo <> "" Then
            If RS.State = 1 Then RS.Close
            RS.Open "select sum(QUANTITY) from ORDERB where INVOICENO=" & txtOrderNo & "", con, adOpenKeyset, adLockOptimistic
            If RS.EOF = False Then
            If tqu >= RS(0) Then
                If txtOrderNo = "" Then Exit Sub
                'If MsgBox("Are you Sure ", vbQuestion + vbYesNo) = vbYes Then
                   con.Execute "update ORDERA set BendingBill='y' where INVOICENO=" & txtOrderNo & ""
                'End If
            End If
            End If
        End If
            
            If RS.State = 1 Then RS.Close
            RS.Open "SELECT top 1 INVOICENO,Scid FROM ApprovalDet where INVOICENO=" & I_NO.text & " group by INVOICENO,Scid", con
            If RS.EOF = False Then
               If RS(1) <> txtScId.text Then
                  con.Execute "delete from ApprovalDet where invoiceno=" & I_NO.text & ""
                  con.Execute "update AppForm set GrossAmt=0,NetAmt=0 where AppNO=" & txtappno.text & ""
                  con.Execute "update invoicea set Appno='',App_Add='n' where invoiceno=" & I_NO.text & ""
                  con.Execute "update invoiceb set Appno='',App_Add='n' where invoiceno=" & I_NO.text & ""
               Else
                  con.Execute "update invoicea set Appno='" & txtappno & "',App_Add='y' where invoiceno=" & I_NO.text & ""
                  con.Execute "update invoiceb set Appno='" & txtappno & "',App_Add='y' where invoiceno=" & I_NO.text & ""
               
               End If
            End If
  
            
            '=========================================
            
            Unload frmEndPartTrans
            Me.customercode.Enabled = False
            Me.Grid1.Enabled = False
            Me.Commandall.Enabled = False
            Me.Commandother.Enabled = False
            Me.Commandadd.Enabled = True
            Me.Commandedit.Enabled = True
            Me.Commandsearch.Enabled = True
            Me.Commandsave.Enabled = False
            Me.Commanddelete.Enabled = True
            Me.Commandabandon.Enabled = True
            Me.CommandPrint.Enabled = True
            Me.CommandDirectPrint.Enabled = True
            Me.Commandprintnh.Enabled = True
        End If
        addmode = False
        addoredit = False
        
        
        mnuMenu_ = "menusalesinvoice"
        SetButton Commandadd, Commandedit, Commandsave, Commanddelete
        
        Check1_trans.Enabled = True
        Check1_withheader.Enabled = True
        
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = False
        Me.Commandadd.SetFocus
        
        If (AuditTrail = "y") Then
        
        If (txtchecked.text = "y") Then
        
            actionType_ = "Edit"
            vtype1_ = "I"
            vtypeNew = "I"
            vdate_ = Trim(i_dt.text)
            vno_ = Trim(I_NO.text)
            
            frmAuditTrailLog_Rem.Show 1
            
         End If
        
        End If
        

Exit Sub
save_:
MsgBox "" & err.Description


End Sub
Private Sub Commandsave_GotFocus()
txtAmtwords = toword(invoice.mna)
End Sub

Private Sub Commandsearch_Click()



sqlQry = "select InvoiceNo,InvoiceDate,Subledger,NetAmount from InvoiceA where " & stringyear & "  InvoiceNo"
orderby = "order by InvoiceNo"


searchType = "inv"
popuplistFast "select InvoiceNo,InvoiceDate,Subledger,NetAmount from InvoiceA where " & stringyear & "  order by InvoiceNo", con, , , "I"



End Sub
Private Sub Commandsearch_GotFocus()
  
If PopUpValue1 <> "" Then

  'If Val(inviceNo) > 0 Then
     I_NO.text = PopUpValue1
     
     PopUpValue1 = ""
     I_NO_LostFocus
     option_withHeader.Enabled = True
     Check1_trans.Enabled = True
     Check1_withheader.Enabled = True
     SetButton Commandadd, Commandedit, Commandsave, Commanddelete
  'End If
  

End If
  
  
End Sub
Private Sub customercode_LostFocus()
    
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "select * from sledger where gledger='SUNDRY DEBTORS' and subledger='" + Trim(customercode.text) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If RS.RecordCount > 0 Then
       lblMail.Caption = RS!email
       lblPAN.Caption = RS!pan & ""
       lblPartyfrt.Caption = RS!freight & ""
       lblPostage.Caption = RS!postage & ""
    End If
    
     Set rs1 = New ADODB.Recordset
     rs1.Open "select top 1 Frt_Yes from ORDERA where INVOICENO ='" & txtOrderNo.text & "'", con, adOpenDynamic, adLockReadOnly, adCmdText
     If rs1.RecordCount > 0 Then
       If Len(rs1!Frt_Yes) > 0 Then
         lblPartyfrt.Caption = rs1!Frt_Yes & ""
       End If
     End If
    
    If (RS.EOF = True Or RS.RecordCount <= 0) Then
        customercode.SetFocus
        HIT
        RS.Close
        Exit Sub
    End If
    
    

    
    Me.textbox.text = Me.customercode.text
    Me.customercode.Visible = False
    Me.customercode.Enabled = False
    
End Sub

Private Sub Delete_Click()
If Grid1.Row >= 1 Then
    Grid1.SetFocus
    Grid1.RemoveItem (Grid1.Row)
    If Grid1.Row > 1 Then
        Grid1.Row = Grid1.Row - 1
    End If
    Grid1_Click
End If
End Sub
Private Sub Form_Activate()

mnuMenu_ = "menusalesinvoice"
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

'    txtMark.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If Grid1.Row >= 1 Then
           Grid1.RemoveItem Grid1.Row
           a = Grid1.text
           tempmeb.text = a
           a = templost
           Grid1.SetFocus
          End If
   End If
End If

If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If UCase(Trim(VB.Screen.ActiveControl.Name)) = UCase(Trim("CUSTOMERCODE")) Then
             If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
                 sendkeys "{tab}"
                 Exit Sub
            End If
             sendkeys "{DOWN}"
             sendkeys "{TAB}"
        Else
            
            If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("grid1")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("tempmeb")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bookname")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("weight")) Then
                sendkeys ("{TAB}")
            End If
        End If
    End If
    
    
    
End Sub

Private Sub Form_Load()
    
Me.top = 0
Me.Left = 0

Me.Width = 14800
Me.Height = 10650



Me.Caption = "Invoice"
    
    
Screen.MousePointer = vbHourglass
    
    
Dim rs_godwn As New ADODB.Recordset

If rs_godwn.State = 1 Then rs_godwn.Close
rs_godwn.Open "select * from GodownMaster where len(Godwn)<=3 and " & stringyear & " order by id", con, adOpenForwardOnly, adLockReadOnly
txtMark.Clear
If Not rs_godwn.EOF Then
Do While Not rs_godwn.EOF
   If IsNull(rs_godwn(0)) = False Then
     Me.txtMark.AddItem rs_godwn(0)
   End If
   If Not rs_godwn.EOF Then rs_godwn.MoveNext
 Loop
End If


If rs_godwn.State = 1 Then rs_godwn.Close
rs_godwn.Open "SELECT DISTINCT Placeofsupply FROM TransportDet order by Placeofsupply", con
While rs_godwn.EOF = False
 txtPlaceSup.AddItem rs_godwn(0)
 rs_godwn.MoveNext
Wend


addmode = False
Edit = False
autoscroll = True
VALIDRATE = True
maxrow = 0
totalamount = 0
totaldiscount = 0
otheramount = 0
otherdiscount = 0
Set kk = New ADODB.Recordset
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Me.top = 50
Me.Left = 50
Grid1.rows = 2
Grid1.Cols = 1
Grid1.rows = 10
Grid1.Cols = 9
Grid1.Row = 0
Grid1.Col = 1
Grid1.text = "Book Code "
Grid1.Col = Grid1.Col + 1
Grid1.text = "Book Name"
Grid1.Col = Grid1.Col + 1
Grid1.text = "Quantity"
Grid1.Col = Grid1.Col + 1
Grid1.text = "Print. Ord."
Grid1.Col = Grid1.Col + 1
Grid1.text = "Rate"
Grid1.Col = Grid1.Col + 1
Grid1.text = "Disc %"
Grid1.Col = Grid1.Col + 1
Grid1.text = "Amount"
Grid1.Col = Grid1.Col + 1
Grid1.text = "Disc. Amount"
Grid1.RowHeight(0) = Grid1.CellHeight + 50
Grid1.ColWidth(0) = 150
Grid1.ColWidth(1) = 1000
Grid1.ColWidth(2) = 3800
Grid1.ColWidth(3) = 1200
Grid1.ColWidth(4) = 1200
Grid1.ColWidth(5) = 1000
Grid1.ColWidth(6) = 1000
Grid1.ColWidth(7) = 1500
Grid1.ColWidth(8) = 1500
Bookname.Height = 2325
Me.CommandPrint.Enabled = True
Me.CommandDirectPrint.Enabled = True
Me.Commandprintnh.Enabled = True


'''RS.Open "select * from books", CON, adOpenDynamic, adLockReadOnly, adCmdText
''Load fron Client
RS.Open "select * from books where " & stringyear, CCON, adOpenDynamic, adLockReadOnly, adCmdText
'Set RS = CON.Execute("exec BookQry '" & session & "'," & main.setupid & "")

If Not RS.BOF Then
    Do While Not RS.EOF
        Me.Bookcode.AddItem RS("bookcode")
        Me.Bookname.AddItem RS("bookname")
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If
RS.Close
    
    
    Genledger.text = "SUNDRY DEBTORS"
    'Set RS = CON.Execute("exec fatch_ledger '" & Genledger.Text & "','" & session & "'," & main.setupid & "")
    RS.Open "select * from sledger where gledger='" & Genledger.text & "' and  " & stringyear, CCON, adOpenDynamic, adLockReadOnly, adCmdText
    
    If Not RS.BOF Then
        Do While Not RS.EOF
            Me.customercode.AddItem RS("subledger")
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    
     '*******Agent  combo fill
    'popuplist10 "select Rep as Representative,Add1,Add2,District,[state] from SalesRepQry order by Rep", CON_blue
    
    'Merge
    
    RS.Open "select Rep as Representative,Email from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
    cmbAgentName.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cmbAgentName.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    
    
    
    RS.Close
    RS.Open "select  transportname from transportMaster order by transportname", con, adOpenDynamic, adLockReadOnly, adCmdText
    cmbtransportname.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cmbtransportname.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    RS.Close

 
    'SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
    On Error Resume Next

    Bookcode.Left = Grid1.Left
    Bookcode.Visible = False
    Bookname.Visible = False
    Grid1.rows = 350
    For I = 1 To 99
        Grid1.RowHeight(I) = 300
    Next
    
    Bookcode.Width = 1230
    Bookname.Width = 2830
    amount.Width = rate.Width

    
    
    If kk.State = 1 Then kk.Close
    kk.Open "SELECT MAX(INVOICENO) FROM INVOICEA where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If kk(0) <> "" Then
       addoredit = False
       
       Me.I_NO.text = kk(0)
       
       If Val(inviceNo) > 0 Then
          Me.I_NO.text = inviceNo
       End If
       
       I_NO_LostFocus
    Else
       Me.I_NO.text = "1"
       i_dt.text = Format(Date, "dd/MM/yyyy")
    End If

    
    

    Commanddelete.Enabled = False
    Commandedit.Enabled = True
    Commandsave.Enabled = False
    lastrow = 0
    lastcol = 1
    
    'Dim ctl As Control
    
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    Picture5.Enabled = True
    

 option_withHeader.Enabled = True
 Check1_trans.Enabled = True
 Check1_withheader.Enabled = True
 Check1_withheader.value = 0
 Check1_trans.value = 0
 Commanddelete.Enabled = False
  
 mnuMenu_ = "menusalesinvoice"
 'SetButton Commandadd, Commandedit, Commandsave, Commanddelete
 
 Commandsave.Enabled = False
 
 Screen.MousePointer = vbDefault
'lblMail
    
 BackColorFrom Me
 Check1_dos.Enabled = True
 Check1_notPrint_inst.Enabled = True
 Check1_spremarks.Enabled = True

End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub freight_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 Then
''        If Trim(Me.customercode.Text) <> "" Then
''            DoEvents
''
''            Grid1.Col = 1
''            Grid1.Row = 1
''            DoEvents
''
''            Grid1_Click
''        Else
''            Me.textbox.SetFocus
''            'Me.customercode.SetFocus
''        End If
''    End If
End Sub

Private Sub freight_LostFocus()

freight = UCase(freight)



End Sub
Private Sub Grid1_Click()

'On Error Resume Next
On Error GoTo ErrorHandler


If Trim(Me.customercode.text) <> "" Then
Dim PREVROW As Integer
Dim prevcol As Integer
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
prevcol = Grid1.Col
PREVROW = Grid1.Row
If Grid1.Row > 1 Then
    Grid1.Row = Grid1.Row - 1
    Grid1.Col = 1
    If Trim(Grid1.text) <> "" Then
        Grid1.Row = PREVROW
        Grid1.Col = prevcol
        If Trim(Me.customercode.text) <> "" Then
            If Me.customercode.Enabled = True Then
                Me.customercode.Enabled = False
            End If
            Grid1.Col = 1
            If prevcol > 1 And Trim(Grid1.text) = "" Then
                Grid1.Col = 2
                sendkeys Chr(13)
            Else
                Grid1.Col = prevcol
                sendkeys Chr(13)
            End If
        Else
            MsgBox "Please fill the customer detail first"
        End If
    End If
Else
    If Trim(Me.customercode.text) <> "" Then
        If Me.customercode.Enabled = True Then
            Me.customercode.Enabled = False
        End If
        Grid1.Col = 1
        If prevcol > 1 And Trim(Grid1.text) = "" Then
            Grid1.Col = 2
            Grid1.SetFocus
            sendkeys Chr(13)
        Else
        'IF GRID1.COL
            Grid1.Col = prevcol
            Grid1.SetFocus
            sendkeys Chr(13)
        End If
        'SendKeys Chr(13)
    End If
End If
End If


   Exit Sub ' Exit before hitting error handler

ErrorHandler:
    ' Handle errors and provide feedback
    MsgBox "An error occurred: " & err.Description, vbCritical, "Error Code: " & err.Number
    Resume Next ' Continue execution after handling the error

End Sub
Private Sub Grid1_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

If Trim(Me.customercode.text) <> "" Then
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
    If (KeyAscii = 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
        If mwritemode = addmode Or mwritemode = EditMode Then
            Dim mprevcol As Integer
            
            Select Case Grid1.Col
            Case 1, 3, 4, 5, 6
                Bookname.Visible = False
                tempmeb.Visible = True: tempmeb.Enabled = True
                tempmeb.ZOrder
                If Grid1.Col <> 1 Then
                    If Grid1.Col <> 3 Then
                        tempmeb.text = Format(Grid1.text, "0.00")
                        
                    Else
                        tempmeb.text = Format(Grid1.text, "0")
                    End If
                   
                Else
                    tempmeb.text = Grid1.text
                End If
                tempmeb.Width = Grid1.ColWidth(Grid1.Col)
                tempmeb.Left = Grid1.CellLeft + Grid1.Left
                tempmeb.top = Grid1.top + Grid1.CellTop
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.text = Grid1.text
                Bookname.top = Grid1.top + Grid1.CellTop
                Bookname.Left = Grid1.CellLeft + Grid1.Left
                Bookname.Width = Grid1.ColWidth(Grid1.Col)
            Case 6
                
            Case Else
                Bookname.Visible = False
                tempmeb.Visible = False
            End Select
            
            Select Case Grid1.Col
                Case 1, 3, 4, 5, 6
                    tempmeb.Mask = ""
                    tempmeb.MaxLength = 20
                Case 2
                    With Bookname
                        .Visible = True
                        .ZOrder
                    End With
             End Select
            
            Select Case Grid1.Col
            Case 2
                Bookname.SetFocus
                If KeyAscii <> 13 Then
                    sendkeys Chr(KeyAscii)
                End If
            Case 1, 3, 4, 5, 6
                mprevcol = Grid1.Col
                tempmeb.SetFocus
            Case Else
                If KeyAscii = 13 Then
                    sendkeys "{RIGHT}"
                End If
            End Select
        End If
    If maxrow < Grid1.Row Then
        maxrow = Grid1.Row
    End If
    
End If
    lastrow = Grid1.Row
    lastcol = Grid1.Col
End If


Exit Sub ' Exit before hitting error handler

ErrorHandler:

    ' Handle errors and provide feedback
    MsgBox "An error occurred: " & err.Description, vbCritical, "Error Code: " & err.Number
    Resume Next ' Continue execution after handling the error

End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

If Button = 2 Then
   PopupMenu dd, , Grid1.Left + X, Grid1.top + Y
End If

   Exit Sub ' Exit before hitting error handler

ErrorHandler:
    ' Handle errors and provide feedback
    MsgBox "An error occurred: " & err.Description, vbCritical, "Error Code: " & err.Number
    Resume Next ' Continue execution after handling the error

End Sub

Private Sub Grid1_Scroll()
If tempmeb.Visible <> True And Bookname.Visible <> True Then
        Bookname.Visible = False
        tempmeb.Visible = False
        'Grid1.SetFocus
End If
autoscroll = True
End Sub

Private Sub i_dt_GotFocus()
'    If Me.Enabled = True Then
'        datedlg.Top = i_dt.Top - 10
'        datedlg.Left = i_dt.Left - 200
'        Me.Enabled = False
'        datedlg.Calendar1.Value = i_dt
'        X = datedlg.GETDATE(Me.Name, i_dt.Name)
'        datedlg.Show
'    End If
If Edit = True Then
    Commandother.Enabled = True
End If
End Sub

Private Sub i_dt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

If Not IsDate(i_dt.text) Then
i_dt.SetFocus
Exit Sub
End If


End If

End Sub
Private Sub i_dt_LostFocus()


If Edit = False Then

If IsDate(i_dt) Then
  If checkData_ForThisNumber("invoicea", I_NO, i_dt) = True Then
      MsgBox "Please select valid Invoice No. for this date.."
      i_dt.SetFocus
  End If
End If

End If

End Sub

Private Sub I_DTOB_LostFocus()
If Trim(I_DTOB.text) <> "__/__/____" Then
    If Not checkdate(Trim(I_DTOB.text), I_DTOB) Then
        I_DTOB.SetFocus
    End If
End If
End Sub

Private Sub I_NO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   If Edit = False Then
   
       If rs1.State = 1 Then rs1.Close
       rs1.Open "select top 1 invoiceno from invoicea where invoiceno=" & I_NO.text & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
       If rs1.EOF = False Then
          MsgBox "This Bill No Alreay Exist...", vbCritical
          I_NO.SetFocus
       End If
     
   End If
End If
End Sub

Private Sub I_NO_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Sub cmdButtonLock()
    Commandother.Enabled = False
    Commandall.Enabled = False
    Commandadd.Enabled = False
    Commandedit.Enabled = False
    Commandsearch.Enabled = False
    Commanddelete.Enabled = False
    Commandabandon.Enabled = False
    Commandprintnh.Enabled = True
    CommandDirectPrint.Enabled = True
    CommandPrint.Enabled = True
End Sub

Sub I_NO_LostFocus()

On Error Resume Next

Dim rs1 As ADODB.Recordset

If Val(inviceNo) > 0 Then
   I_NO.text = inviceNo
   cmdButtonLock
End If



inviceNo = ""

Set rs1 = New ADODB.Recordset
Set RS = New ADODB.Recordset

    If Trim(I_NO.text) = "" Then
        MsgBox "Invoice cannot be null"
        I_NO.SetFocus
    Else
        
        If RS.State = 1 Then
           RS.Close
        End If
        RS.Open "Select top 1 * from  INVOICEA where INVOICENO = " + Trim(I_NO.text) + " and " & stringyear, con, adOpenStatic, adLockReadOnly
        If RS.EOF Then
            If addoredit = False Then
                MsgBox "Invoice not found"
                Exit Sub
            End If
            Exit Sub
        End If
        If addoredit Then
            MsgBox "Invoice already exist..."
            If I_NO.Enabled = False Then Exit Sub
            I_NO.SetFocus
            HIT
            Exit Sub
        End If
        
        
       invoiceabandon

        
        If Not IsNull(RS!todid) Or RS!todid = "" Then
           txtTODNO.text = RS!todid
           txtTODDate.text = RS!toddate
        End If

        
        If (RS!Shipto_Scholl = 1) Then
           Check1_school.value = 1
        Else
           Check1_school.value = 0
        End If
        
        
        If (AuditTrail = "y") Then
        
            If (RS!Checked_YesNo = True) Then
               txtchecked.text = "y"
            Else
                txtchecked.text = "n"
            End If
        
        End If
        
        If RS!SMSDate <> "" Then
          txtRandomDT.text = RS!SMSDate
        End If
        txtRandomId.text = RS!randomId & ""
        txtRandomMob.text = RS!mobile
        
        txtPendingBooksRem.text = RS!PendingRemarks & ""
        
        lblApp(4).Caption = RS!app_add & ""
        txtappno.text = RS!appno & ""
         
        txtcreatedDT.text = RS!CreatedDt
         
        LblRandomNo.Caption = RS!RandomNo
        txtRem = RS!remarks & ""
        txtPlaceSup = RS!Placeofsupply & ""
        txtNSCHNo.text = RS!NsChallanNo & ""
        
        txtschool = RS!scname & ""
        txtScId.text = RS!scid & ""
        
        If (IsNull(RS!Shipto_City) Or RS!Shipto_City = "") Then
           txtShip = RS!Shipto
        Else
           txtShip = RS!Shipto & "," & RS!Shipto_City & ""
        End If
        
        lblBookSId.Caption = RS!Shipto_CityId & ""

        I_NO.text = RS!invoiceNo
        Me.i_dt.text = RS!invoiceDate
        Me.Genledger.text = Trim(RS!Genledger)
        Me.customercode.text = Trim(RS!subledger)
        Me.cmbAgentName.text = IIf(IsNull(RS!agentname), "", RS!agentname)
        Me.cmbtransportname.text = IIf(IsNull(RS!transportname), "", RS!transportname)
        Me.textbox.text = Trim(RS!subledger)
        Me.I_OB.text = IIf(IsNull(RS!orderby), "", Trim(RS!orderby))
        If RS!ORDERDATE <> "" Then
        Me.I_DTOB.text = RS!ORDERDATE
        End If
        txtOrderNo.text = RS!orderNo & ""
        
        Me.marka.text = IIf(IsNull(RS!marka), "", Trim(RS!marka))
        txtMark.text = RS!Godown & ""
        Me.bundles = IIf(IsNull(RS!bundles), "", RS!bundles)
        Me.through.text = IIf(IsNull(RS!through), "", RS!through)
        Me.through1.text = IIf(IsNull(RS!through1), "", RS!through1)
        Me.station.text = IIf(IsNull(RS!station), "", RS!station)
        Me.biltno.text = IIf(IsNull(RS!biltyno), "", RS!biltyno)
        If RS!BILTYDATE <> "" Then
        Me.bdated = RS!BILTYDATE
        End If
        Me.freight = IIf(IsNull(RS!freight), "", RS!freight)
        Me.weight = IIf(IsNull(RS!weight), "", RS!weight)
        'Me.labelbybank = round(val(Trim(rs!baa)
        Me.labelbybank = Format(Round(RS!baa, 2), "0.00")
       ' mna.Caption = rs!netamount
        mna.Caption = Format(Round(RS!netamount, 2), "0.00")
        'Me.Combosldistrictcode.Text = IIf(IsNull(rs!district), "", rs!district)
        Me.txtadst = IIf(IsNull(RS!advicestatus), "", RS!advicestatus)
        
        txtAmtwords = RS!Amtwords & ""
        lblSMSId.Caption = RS!randomId & ""
         
        '=====================================
        If RS.State = 1 Then RS.Close
        RS.Open "select pan from sledger where gledger='SUNDRY DEBTORS' and subledger='" + Trim(Me.customercode.text) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.RecordCount > 0 Then
           lblPAN.Caption = RS!pan & ""
        End If
        
        '=====================================
       
       ' OTHERSALES.Form_Load
       '*/**/*/*/*/*//*/*
       
        If RS.State = 1 Then RS.Close
        
       con.Execute "select * from INVOICEctmp WHERE INVOICENO=" & invoice.I_NO & " and " & stringyear
       RS.Open "Select * from INVOICEB where INVOICENO =" + Trim(I_NO.text) + " and " & stringyear & " order by SNO", con, adOpenStatic, adLockReadOnly
       Grid1.TopRow = 2
        If Not RS.EOF Then
        
            Grid1.Row = 1
            Grid1.Col = 1
            Do While Not RS.EOF
            aa = RS.RecordCount
               If Trim(RS!invoiceNo) = Trim(I_NO.text) Then
                Grid1.Col = 1
                Grid1.text = Trim(RS!Bookcode)
                If kk.State = 1 Then
                    kk.Close
                End If
                kk.Open "select * from books where bookcode='" + Trim(RS!Bookcode) + "' and " & stringyear & "", con, adOpenStatic, adLockReadOnly, adCmdText
                Grid1.Col = 2
                Grid1.text = Trim(kk!Bookname)
                Grid1.Col = 3
                Grid1.text = Trim(RS!QUANTITY)
                Grid1.Col = 5
                Grid1.text = Format(Round(RS!rate, 2), "0.00")
                Grid1.Col = 7
                Grid1.text = Format(Round(RS!amount, 2), "0.00")
                Grid1.Col = 4
                Grid1.text = Format(Round(RS!PRINTORDER, 2), "0.00")
                Grid1.Col = 6
                Grid1.text = Format(Round(RS!discount, 2), "0.00")
                Grid1.Col = 8
                Grid1.text = Format(Round(RS!amount * (RS!discount / 100), 2), "0.00")
                End If
                If Not RS.EOF Then
                    RS.MoveNext
                End If
                Grid1.Row = Grid1.Row + 1
                Grid1.rows = Grid1.rows + 1
            Loop
            maxrow = Grid1.Row
            'Me.i_dt.SetFocus
            
        End If
        Row = Grid1.Row
        Col = Grid1.Col
        Grid1.TopRow = 1
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Round(Val(Trim(Grid1.text)), 2)
            Grid1.Col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.text)), 2)
        Next
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     Me.tqu.Caption = ""
        For I = 1 To maxrow
            Grid1.Col = 3
            Grid1.Row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.text))
        Next
        Grid1.Row = RRR
        Grid1.Col = CCC
       ' templost = True
    End If
    Me.Commandother.Enabled = True
    
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select mail,freight_yes_no,postage_yes_no from invoiceaQry where INVOICENO=" & I_NO & "", con
    If rs1.EOF = False Then
       lblMail.Caption = rs1!mail & ""
       lblPartyfrt.Caption = rs1!freight_yes_no & ""
       lblPostage.Caption = rs1!postage_yes_no & ""
    End If
         
    
    
    
    
     If Val(inviceNo) > 0 Then
      I_NO.text = inviceNo
      cmdButtonLock
   End If
   
   DoEvents
   DoEvents
   DoEvents
   DoEvents
   Grid1.Redraw = True
   
   Commandother.Enabled = False
   Commanddelete.Enabled = False
   Check1_spremarks.Enabled = True
   
   mnuMenu_ = "menusalesinvoice"
   ''SetButton Commandadd, Commandedit, Commandsave, Commanddelete
   
End Sub

Private Sub I_OB_LostFocus()
I_OB = UCase(I_OB)
End Sub


Private Sub marka_GotFocus()
Dim trs As New ADODB.Recordset
 
End Sub

Private Sub marka_LostFocus()
marka = UCase(marka)


End Sub


Private Sub station_LostFocus()
station = UCase(station)
End Sub
Private Sub tempmeb_Change()

On Error GoTo ErrorHandler

If Grid1.Col = 1 Or Grid1.Col = 2 Then
    Grid1.text = tempmeb.text
Else
    If Grid1.Col = 3 Then
        Grid1.text = Format(tempmeb.text, "0")
    Else
        Grid1.text = Format(tempmeb.text, "0.00")
    End If
End If

   Exit Sub ' Exit before hitting error handler

ErrorHandler:
    ' Handle errors and provide feedback
    MsgBox "An error occurred: " & err.Description, vbCritical, "Error Code: " & err.Number
    Resume Next ' Continue execution after handling the error

End Sub
Private Sub tempmeb_GotFocus()
    HIT
     
End Sub
Private Sub tempmeb_KeyPress(KeyAscii As Integer)

On Error GoTo aa11

If KeyAscii = 13 Then

        Dim RS As ADODB.Recordset
           Set RS = New ADODB.Recordset
            Select Case Grid1.Col
                Case 1
                    'If RS.State = 1 Then RS.close
                    Set RS = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(Grid1.text) & "'")
                    'RS.Open "books", CON, adOpenStatic, adLockReadOnly, adCmdTable
                    If RS.BOF = True Then
                       Grid1.Col = 2
                       'Exit Sub
                    End If
                    If Not RS.BOF Then
                        'RS.MoveFirst
                        'RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF And Trim(Grid1.text) <> "" Then
                            RS.Close
                            Exit Sub
                        Else
                            RS.Close
                        If Trim(Grid1.text) <> "" Then
                                Grid1.Col = 3
                            Else
                                Grid1.Col = 2
                            End If
                        End If
                    Else
                        If Trim(Grid1.text) <> "" Then
                            Grid1.Col = 3
                        Else
                            Grid1.Col = 2
                        End If
                    End If
                    Grid1.SetFocus
                    Grid1_Click
                    
                Case 3
                
                    If Val(tempmeb.text) > 0 Then
                    
                        Grid1.Col = 1
                        'grid1.Col = grid1.Col + 2
                        Grid1.Row = Grid1.Row + 1
                        Grid1.rows = Grid1.rows + 1
 
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                Case 4
                    Grid1.Col = Grid1.Col + 2
                    Grid1.SetFocus
                    Grid1_Click
                Case 5
                    If Val(tempmeb.text) > 0 Then
                        Grid1.Col = Grid1.Col - 1
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                Case 6
                
                   If Val(Grid1.TextMatrix(Grid1.Row, 4)) <> Val(Grid1.TextMatrix(Grid1.Row, 6)) Then
                      MsgBox "Discount And Printorder Not Match.."
                      
                   End If
                    Grid1.Col = 1
                    Grid1.Row = Grid1.Row + 1
                    Grid1.rows = Grid1.rows + 1
                    Grid1.SetFocus
                    Grid1_Click
            End Select
        Else
        If Grid1.Col = 3 Or Grid1.Col = 4 Or Grid1.Col = 5 Or Grid1.Col = 6 Then
            If KeyAscii >= 48 And KeyAscii <= 57 Then
            Else
                If KeyAscii <> 46 Then
                    If KeyAscii <> 8 Then
                        KeyAscii = 0
                    End If
                End If
            End If
        End If
    End If
    


Exit Sub ' Exit before hitting error handler

aa11:
' Handle errors and provide feedback
MsgBox "An error occurred: " & err.Description, vbCritical, "Error Code: " & err.Number
Resume Next ' Continue execution after handling the error
    
End Sub
Private Sub tempmeb_LostFocus()
    
On Error GoTo ErrorHandler
    
    If templost Then
        tempmeb.Visible = False
    End If

   Exit Sub ' Exit before hitting error handler

ErrorHandler:
    ' Handle errors and provide feedback
    MsgBox "An error occurred: " & err.Description, vbCritical, "Error Code: " & err.Number
    Resume Next ' Continue execution after handling the error

End Sub
Private Sub textbox_GotFocus()
    
Me.customercode.Enabled = True
Me.customercode.Visible = True
'Me.customercode.Height = 1100
Me.customercode.ZOrder
Me.customercode.SetFocus
    
End Sub
Private Sub through_LostFocus()
through = UCase(through)
End Sub
Private Sub through1_LostFocus()
through1 = UCase(through1)
End Sub
Private Sub txtOrderNo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   
   
If Edit = False Then

If Val(txtOrderNo) > 0 Then

   If txtOrderNo <> "" Then
   PopUpValue1 = txtOrderNo
   End If
   
   invoiceabandon
   fatchOrder
   On Error Resume Next
    Dim ctl As Control
    For Each ctl In Me.Controls
    If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
        If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) Then
        End If
        ctl.Enabled = True
    End If
    Next
    
End If
   
End If
    
I_DTOB.SetFocus
    
End If






End Sub
Private Sub txtOrderNo_LostFocus()

If Len(txtOrderNo) > 0 Then
    If Not IsNumeric(txtOrderNo) Then
       MsgBox "This OrderNo is Numeric value....", vbCritical
       txtOrderNo.SetFocus
    End If
End If

End Sub

Private Sub txtPendingBooksRem_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
'If MsgBox("Are to Edit ?", vbYesNo) = vbYes Then
 
'End If
txtPendingBooksRem.text = UCase(txtPendingBooksRem.text)
End If

End Sub
Private Sub txtPendingBooksRem_LostFocus()
 con.Execute ("update invoicea set PendingRemarks='" & Trim(txtPendingBooksRem.text) & "' where INVOICENO = " + Trim(I_NO.text))
End Sub
Private Sub txtPlaceSup_LostFocus()
 txtPlaceSup.text = UCase(txtPlaceSup)
End Sub

Private Sub txtRem_LostFocus()
txtRem = UCase(txtRem)
End Sub

Private Sub txtschool_GotFocus()

'If RS.State = 1 Then RS.close
If PopUpValue1 <> "" Then
'
'txtScId = PopUpValue1
'txtSchool.text = PopUpValue2 & ", " & PopUpValue3


txtschool.text = PopUpValue1
txtScId.text = PopUpValue2
  


PopUpValue1 = ""
PopUpValue2 = ""
'PopUpValue3 = ""
'popupvalue4 = ""
'
End If




End Sub

Private Sub txtschool_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then

'   Screen.MousePointer = vbHourglass
'   tblNo = 9
'   frmSearchItem.Show
'   Screen.MousePointer = vbDefault



searchType = "party"
popuplist_client "SELECT distinct ScName,ScID FROM ORDERA where partyname ='" & textbox.text & "' order by scname", con


End If





End Sub
Private Sub txtschool_LostFocus()
If txtschool.text = "" Then
   txtScId.text = ""
End If
End Sub

Private Sub txtShip_GotFocus()

If Check1_school.value = 0 Then

If RS.State = 1 Then RS.Close
If PopUpValue1 <> "" Then
    txtShip = PopUpValue2 & "," & popupvalue4
    lblBookSId.Caption = PopUpValue1
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""
End If


Else
    'for school--------------------
    '------------------------------
    
    If RS.State = 1 Then RS.Close
    If PopUpValue1 <> "" Then
        txtShip = PopUpValue2 & "," & popupvalue4
        lblBookSId.Caption = PopUpValue1
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        popupvalue4 = ""
    End If

End If

End Sub

Private Sub txtShip_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    If Check1_school.value = 0 Then
      
      tblNo = 51
      frmSearchItem.Show
      
    Else
       '''For School
      Screen.MousePointer = vbHourglass
      tblNo = 9
      frmSearchItem.Show
      Screen.MousePointer = vbDefault
    End If
End If

End Sub
Private Sub txtShip_LostFocus()
   If Len(txtShip) = 0 Then
      lblBookSId.Caption = ""
   End If
   txtShip = UCase(txtShip)
End Sub

Private Sub weight_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then freight.SetFocus
End Sub

Private Sub weight_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        If Trim(Me.customercode.text) <> "" Then
            Grid1.Col = 1
            Grid1.Row = 1
            Grid1_Click
        Else
            Me.textbox.SetFocus
            'Me.customercode.SetFocus
        End If
    End If

End Sub

Private Sub weight_LostFocus()
weight = UCase(weight)



Dim amt, total_ As Double
amt = 0

If (cmbtransportname.text <> "" And txtPlaceSup.text <> "") Then
  Set rs1 = New ADODB.Recordset
  rs1.Open "SELECT GeneralRate, Doordelivery from TransportDet where (TransportName='" & cmbtransportname.text & "' and Placeofsupply='" & txtPlaceSup.text & "')", con
  If rs1.EOF = False Then
     If rs1!GeneralRate > 0 Then
        amt = rs1!GeneralRate
     Else
        amt = rs1!Doordelivery
     End If
     
     
     
     If amt > 0 Then
        If Val(freight.text) > 0 Then
           total_ = (Val(weight) * amt + 100)
           If Val(freight.text) > total_ Then
              DoEvents
              DoEvents
              MsgBox "Freight is Greater than given rate..." & vbCrLf & amt & " Rate/kg " & "", vbCritical
              
              DoEvents
              Grid1.Col = 1
              Grid1.Row = 1
              Me.tempmeb.SetFocus
              DoEvents
              
           End If
        End If
     End If
     
     
     
     
  End If
  
End If

End Sub
Sub tt()

Dim flagyes As Boolean
   flagyes = True
If MsgBox("Print Head Yes/No", vbYesNo) = vbNo Then

    flagyes = False

End If
 
 
 
 
 Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = False
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = False
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
        Me.CommandDirectPrint.Enabled = True
        Me.Commandprintnh.Enabled = True
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim netamount As Double
    Dim totalquantity As Long
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    T1 = 10
    T2 = 25
    T3 = 40
    T4 = 55
    T5 = 70
    T6 = 85
    T7 = 100
    T8 = 115
    netamount = 0
    totalquantity = 0
    paperWidth = 150
    MaxLine = 70
    called1 = False
    called2 = False
    Dim Line As Integer
    Dim rs1 As ADODB.Recordset
    Dim kkk As ADODB.Recordset
    Set kkk = New ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    
    Open "" + VB.App.Path + "\vipin.txt" For Output As #1
    Line = 0
header:
      If kkk.State = 1 Then
            kkk.Close
      End If
      If flagyes = True Then
      CNSetup
          kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
          If Not kkk.BOF Then
             Print #1, Chr(27) + Chr(18) + Chr(14)
             Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(18) + Chr(14); Trim(kkk!cname)
             Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(18); dspace(Trim(kkk!add1))
             Print #1, Tab((paperWidth - Len(Trim(kkk!phone1))) / 2); Trim(kkk!phone1)
             Line = Line + 3
         End If
          If rs1.State = 1 Then
             rs1.Close
          End If
          Print #1, Chr(27) + Chr(14)
          Line = Line + 1
      
    Else
       Print #1, ""
       Print #1, ""
       Print #1, ""
       Print #1, Chr(27) + Chr(18)
       Line = Line + 4
  End If
  
  
  
    If rs1.State = 1 Then
        rs1.Close
    End If
    rs1.Open "invoicea", con, adOpenDynamic, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.text) + "'"
   
    Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!subledger; Tab(T5); "Invoice No. : "; Trim(rs1!invoiceNo); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!invoiceDate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.Close
            End If
            kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE;
                Print #1, Tab(3); kkk!address1; Tab(T5); "Order by    : "; Trim(rs1!orderby); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!ORDERDATE
                Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.   : "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!BILTYDATE
                
                
                kkk.Close
                Print #1,
                Print #1, "Through  :"; Tab(12); Trim(rs1!through) + ", " + Trim(rs1!through1)
                Print #1, "Station  :"; Tab(12); Trim(rs1!station); Tab(T5); "Pvt. Mark   : "; Trim(rs1!marka)
                Print #1, "Freight  :"; Tab(12); Trim(rs1!freight); Tab(T5); "Weight      : "; Trim(rs1!weight); Tab(T7 + 7); "Bundle(s)   : "; Trim(rs1!bundles)
                Print #1, repli("-", 150)
                Print #1, "S.No."; Tab(11); "Book Description"; Tab(T5 - 3); "Quantity"; Tab(T6 + 4); "Rate"; Tab(T7 + 4); "Amount"; Tab(T8 + 9); "Net Amount"
                Print #1, repli("-", 150)
                Line = Line + 11
            End If
            If called1 Then
                GoTo printagain1
            End If
            If called2 Then
                GoTo printagain2
            End If
            If kk.State = 1 Then
                kk.Close
            End If
            kk.Open "select * from INVOICEB where invoiceno=" + Trim(rs1!invoiceNo) + " and " + stringyear + " order by discount,printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kk.BOF Then
                kk.MoveFirst
                Dim cdiscount As Double
                Dim sno As Integer
                Dim tdata As ADODB.Recordset
                Set tdata = New ADODB.Recordset
                sno = 1
                Do While Not kk.EOF
                    cdiscount = kk!discount
                    Do While kk!discount = cdiscount
                        tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
                        Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(Str(kk!QUANTITY)), 5); Tab(T6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                        totalquantity = totalquantity + kk!QUANTITY
                        Line = Line + 1
                        If Line > MaxLine Then
                            called1 = True
                            Line = 0
                            Print #1, Chr(12)
                            GoTo header
printagain1:
                            called1 = False
                        End If
                        tdata.Close
                        If Not kk.EOF Then
                            sno = sno + 1
                            kk.MoveNext
                        End If
                        If kk.EOF Then
                            Exit Do
                        End If
                    Loop
                        If Line > MaxLine - 4 Then
                            called2 = True
                            Line = 0
                            Print #1, Chr(12)
                            GoTo header
printagain2:
                            called2 = False
                        End If
                        
                        Print #1, Tab(T7); repli("-", 22)
                        tdata.Open "select sum(amount) from INVOICEB where invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " and " & stringyear & " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
                        If Not tdata.BOF Then
                            
                            Print #1, Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0), 2)), "0.00")), 12)
                            Print #1, Tab(T5); "Less Discount @ " + Trim(Format(Str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(Str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                            netamount = netamount + Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                        End If
                        tdata.Close
                        Print #1, Tab(T7); repli("-", 22)
                Loop
            End If
           End If
           Print #1, Tab(T5 - 4); rsets(Trim(Str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
           Print #1, Tab(T6); repli("-", 22)
           If kk.State = 1 Then
                kk.Close
           End If
           kk.Open "Select * from invoicec where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
           If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5); Trim(kk!text) + "    " + Trim(Format(Str(Round(kk!rate, 2)), "0.00")); Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5); Trim(kk!text); Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T6); repli("-", 22)
            Print #1, Tab(6); Chr(71) + "NET AMOUNT: "; Tab(T8 + 5); Chr(72) + rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
           End If
           kk.Close
           kk.Open "Select * from invoicea where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
           If Not kk.BOF Then
                If kk!txt1a <> 0 Then
                    Print #1, Tab(T5); kk!txt1; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                    netamount = netamount + Round(kk!txt1a, 2)
                End If
                If kk!txt2a <> 0 Then
                    Print #1, Tab(T5); kk!txt2; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                    netamount = netamount + Round(kk!txt2a, 2)
                End If
                If kk!baa <> 0 Then
                    Print #1, Tab(T5); "BY BANK "; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                    netamount = netamount - Round(kk!baa, 2)
                End If
           End If
           Print #1, Tab(T6); repli("-", 22)
           Print #1, Tab(T5); Chr(71) + "BALANCE : "; Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12) + Chr(72)
        
       ' PRINT THE FOOTER IN INVOICE START
       
            Do While Line < MaxLine
                    Print #1, " "
                    Line = Line + 1
            Loop
       
       
            Print #1, Tab(0); repli("-", 120)
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            Dim LEFTM As Integer
            LEFTM = 5
            CNSetup
            tempdata.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
            Print #1, Tab(1); "E.& O.E"
            Print #1, Tab(LEFTM); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!cname)) * 0.75))); "FOR " + Trim(tempdata!cname)
            Print #1, Tab(LEFTM); ""

       'PRINT THE FOOTER IN INVOICE END
       
       
        
        
        
        
        
        Close #1
         PrintOption.Show
        'Me.Enabled = False
        'viewinvoice.Left = 0
        'viewinvoice.Top = 10
        'viewinvoice.Show



End Sub














'***888888888888888888**************************************************************



Sub Bkupprintinvoice()

Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = False
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = False
Me.Commandabandon.Enabled = True
Me.CommandPrint.Enabled = True
Me.CommandDirectPrint.Enabled = True
Me.Commandprintnh.Enabled = True
Dim called1, called2 As Boolean
Dim MaxLine As Integer
Dim netamount As Double
Dim totalquantity As Long
Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
Dim RS As ADODB.Recordset
Dim Pno As Integer
Set RS = New ADODB.Recordset
T1 = 10
T2 = 25
T3 = 40
T4 = 55
T5 = 70
T6 = 85
T7 = 100
T8 = 115
netamount = 0
totalquantity = 0
paperWidth = 145
MaxLine = 60
called1 = False
called2 = False

Dim Line As Integer
Dim rs1 As ADODB.Recordset
Dim kkk As ADODB.Recordset
Dim FooterYes As Boolean
Set kkk = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim LEFTM As Integer
Open "" + VB.App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
FooterYes = False
header:
If kkk.State = 1 Then
      kkk.Close
End If
CNSetup
kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
If FooterYes = True Then
   If Line > MaxLine - 4 Then
       Do While Line < 61
            Print #1, ""
            Line = Line + 1
       Loop
   End If
   Line = 0
   LEFTM = 5
   Print #1, Tab(0); repli("-", 145)
   Print #1, Tab(1); "E.& O.E"
   Print #1, Tab(1); kkk!COURT; Tab(LEFTM + (paperWidth - ((Len(kkk!COURT) + Len(kkk!cname)) * 0.75))); "FOR " + Trim(kkk!cname)
   Print #1, ""
   Print #1, Tab(1); "Continued on Page : " & Pno
   Print #1, ""
   Print #1, ""
   Print #1, ""
   Print #1, ""
   Print #1, ""
   Print #1, ""
   Print #1, ""
End If
If Printheader = True Then
   If Not kkk.BOF Then
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, Chr(27) + Chr(71); Chr(27) + Chr(15) + Chr(14)
     Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname)
     Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
     Print #1, Tab((paperWidth - Len(Trim(kkk!phone1))) / 2); Trim(kkk!phone1) & "," & Trim(kkk!phone2)
     Line = Line + 9
   End If
Else
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Line = Line + 8
End If
Print #1, Tab(1); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(15) + Chr(14); Tab(30); dspace(Trim("INVOICE")); Chr(20); Tab(T4 + 6); IIf(Printheader = True, kkk!uptt, "")
If Printheader = True Then
   Print #1, Tab(T7 + 4); kkk!cst
Else
   Line = Line - 1
End If

Print #1, repli("-", 145)
Line = Line + 3
If rs1.State = 1 Then
   rs1.Close
End If
'Print #1, Chr(27) + Chr(14)
'line = line + 1
If rs1.State = 1 Then
    rs1.Close
End If
rs1.Open "invoicea", con, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.text) + "'"
If Not rs1.EOF Then
    Print #1, " To, S.L. Code : "; Tab(T1 + 10); Mid$(rs1!subledger, 1, 5); Tab(T5); "Invoice No. : "; Trim(rs1!invoiceNo); Tab(T8); "Dated     : "; rs1!invoiceDate   'Chr(27) + Chr(18);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE
        Print #1, Tab(3); kkk!address1; Tab(T5); "Order by    : "; Trim(rs1!orderby); Tab(T8); "Dated     : "; IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
        Print #1, Tab(3); kkk!address2; Tab(T5); "Bilty No.   : "; Trim(rs1!biltyno); Tab(T8); "Dated     : "; IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, Tab(3); kkk!address3
        kkk.Close
        Print #1, "Through  :"; Tab(12); Trim(rs1!through) + IIf(Trim(rs1!through1) = "", "", "," & rs1!through1)
        Print #1, "Station  :"; Tab(12); Trim(rs1!station); Tab(T5); "Pvt. Mark   : "; Trim(rs1!marka)
        Print #1, "Freight  :"; Tab(12); Trim(rs1!freight); Tab(T5); "Weight      : "; Trim(rs1!weight); Tab(T8); "Bundle(s) : "; Trim(rs1!bundles); Chr(27) + Chr(72)
        Line = Line + 1
        Print #1, repli("-", 145)
        Print #1, "S.No."; Tab(11); "Book Description"; Tab(T5 - 3); "Quantity"; Tab(T6 + 4); "Rate"; Tab(T7 + 4); "Amount"; Tab(T8 + 9); "Net Amount"
        Print #1, repli("-", 145)
        Line = Line + 10
    End If
    If called1 Then
        GoTo printagain1
    End If
    If called2 Then
        GoTo printagain2
    End If
    If kk.State = 1 Then
        kk.Close
    End If
    kk.Open "select * from INVOICEB where invoiceno=" + Trim(rs1!invoiceNo) + " and " & stringyear & " order by discount,printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
        kk.MoveFirst
        Dim cdiscount As Double
        Dim sno As Integer
        Dim tdata As ADODB.Recordset
        Set tdata = New ADODB.Recordset
        sno = 1
        Do While Not kk.EOF
            cdiscount = kk!discount
            Do While kk!discount = cdiscount
                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
                Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(Str(kk!QUANTITY)), 5); Tab(T6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!QUANTITY
                Line = Line + 1
                If Line > MaxLine - 4 Then
                    called1 = True
                    'Line = 0
                    'Print #1, Chr(12)
                    'Line = Line + 1
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
printagain1:
                    called1 = False
               End If
                tdata.Close
                If Not kk.EOF Then
                    sno = sno + 1
                    kk.MoveNext
                End If
                If kk.EOF Then
                    Exit Do
                End If
            Loop
                If Line > MaxLine - 4 Then
                    called2 = True
                    'Line = 0
                    'Print #1, Chr(12)
                    'Line = Line + 1
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
                    
'''                    Do While Line < 63
'''                       Print #1, ""
'''                       Line = Line + 1
'''                    Loop
'''
'''
'''                    GoTo foot
                    
printagain2:
                    
                    called2 = False
                End If
                Print #1, Tab(T7); repli("-", 22)
                Line = Line + 1
                tdata.Open "select sum(amount) from INVOICEB where invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " and " & stringyear & " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                    Print #1, Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0), 2)), "0.00")), 12)
                    Print #1, Tab(T5); "Less Discount @ " + Trim(Format(Str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(Str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                    Print #1, Tab(T7); repli("-", 22)
                    Line = Line + 3
                    netamount = netamount + Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                End If
                tdata.Close
                'Print #1, Tab(t7); repli("-", 22)
                'line = line + 1
                Loop
            End If
        End If
        Print #1, repli("-", 145)
        Print #1, Tab(T5 - 4); rsets(Trim(Str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
       'by vk Print #1, Tab(t8); repli("-", 22)
        Line = Line + 2
        If kk.State = 1 Then
             kk.Close
        End If
        kk.Open "Select * from invoicec where " & stringyear & " and  invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5 + 21); Trim(kk!text) + " :  @  " + Trim(Format(Str(Round(kk!rate, 2)), "0.00")) & " % "; Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5 + 20); Trim(kk!text) & " :"; Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T8); repli("-", 22)
            Print #1, Chr(27) + Chr(71); Tab(T5 + 20); "NET AMOUNT  : "; Tab(T8 + 6); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
            VNetamt = netamount
        Line = Line + 2
        End If
        kk.Close
        kk.Open "Select * from invoicea where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(T5 + 20); kk!txt1 & "  :"; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + Round(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(T5 + 20); kk!txt2 & " :"; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - Round(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(T5 + 20); "BY BANK       :"; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - Round(kk!baa, 2)
             End If
        End If
        Print #1, Tab(T8); repli("-", 22)
        Print #1, Tab(T5 + 20); Chr(27) + Chr(71); "BALANCE    : "; Tab(T8 + 6); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
        Print #1, Tab(T8); repli("-", 22)
        Line = Line + 3
       ' PRINT THE FOOTER IN INVOICE START
        Do While Line < 61
            Print #1, ""
            Line = Line + 1
        Loop
        Print #1, Tab(0); toword(Round(VNetamt, 2))

        Print #1, Tab(0); repli("-", 145)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        'Dim LEFTM As Integer
        'LEFTM = 5
        CNSetup
        tempdata.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!cname)) * 0.75))); "FOR " + Trim(tempdata!cname)
    
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        'PRINT THE FOOTER IN INVOICE END
        
        Close #1
        PrintOption.Show
        
        'Me.Enabled = False
        'viewinvoice.Left = 0
        'viewinvoice.Top = 10
        'viewinvoice.Show

End Sub

Function rsets(ST As String, length As Integer) As String
   
    Dim kk As String
            kk = Trim(ST)
            If Len(kk) < length Then
                Do While Not Len(kk) = length
                    kk = " " + kk
                Loop
            End If
            If Len(kk) > length Then
                Do While Not Len(kk) = length
                    kk = Mid$(kk, 0, Len(kk) - 1)
                Loop
            End If
        rsets = kk
End Function



