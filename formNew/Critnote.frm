VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Critnote 
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13992
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10440
   ScaleWidth      =   13992
   Begin VB.CheckBox Check_header 
      Caption         =   "Print With Header"
      Height          =   195
      Left            =   5940
      TabIndex        =   73
      Top             =   9864
      Width           =   1770
   End
   Begin VB.Frame panel 
      Caption         =   "Credit Note Item"
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
      Height          =   10260
      Left            =   60
      TabIndex        =   17
      Top             =   120
      Width           =   13920
      Begin VB.TextBox txtchecked 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   4620
         MaxLength       =   100
         TabIndex        =   77
         Top             =   8460
         Width           =   540
      End
      Begin VB.CheckBox Check1_direct 
         Caption         =   "Direct Print"
         Height          =   195
         Left            =   9216
         TabIndex        =   76
         Top             =   9864
         Width           =   1395
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
         Height          =   5544
         Left            =   540
         TabIndex        =   75
         Top             =   4104
         Visible         =   0   'False
         Width           =   1848
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
         Left            =   324
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   9828
         Width           =   1092
      End
      Begin VB.CheckBox Check1_dos 
         Caption         =   "Show Screen  "
         Height          =   195
         Left            =   7812
         TabIndex        =   70
         Top             =   9864
         Width           =   1395
      End
      Begin VB.TextBox txtTODNO 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   66
         Top             =   8460
         Width           =   1035
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   64
         Top             =   8100
         Width           =   7875
      End
      Begin VB.TextBox txtAmtwords 
         Appearance      =   0  'Flat
         Height          =   540
         Left            =   9720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   8535
         Width           =   3855
      End
      Begin VB.TextBox txtScId 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4905
         TabIndex        =   59
         Top             =   1200
         Width           =   570
      End
      Begin VB.ComboBox cmbAgentName 
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   1500
         Width           =   2025
      End
      Begin VB.CommandButton Commandall 
         Caption         =   "All Books"
         Height          =   600
         Left            =   1440
         TabIndex        =   33
         Top             =   7260
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton Commandother 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&End Part"
         Height          =   600
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7275
         Width           =   930
      End
      Begin VB.ComboBox Bookname 
         Height          =   912
         Left            =   3705
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   32
         Top             =   3090
         Width           =   2295
      End
      Begin VB.ComboBox Bookcode 
         Height          =   1872
         Left            =   735
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   31
         Top             =   2580
         Width           =   2355
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   315
         ScaleHeight     =   840
         ScaleWidth      =   9288
         TabIndex        =   19
         Top             =   8856
         Width           =   9285
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   720
            Left            =   6840
            Picture         =   "Critnote.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   90
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N&HPrint"
            Enabled         =   0   'False
            Height          =   720
            Left            =   7890
            Picture         =   "Critnote.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   90
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.CommandButton Commandadd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   720
            Left            =   30
            Picture         =   "Critnote.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   90
            Width           =   1095
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   720
            Left            =   8070
            Picture         =   "Critnote.frx":23AC
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   90
            Width           =   1155
         End
         Begin VB.CommandButton CommandPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print(det)"
            Enabled         =   0   'False
            Height          =   720
            Left            =   6780
            Picture         =   "Critnote.frx":2F90
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   90
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   720
            Left            =   5580
            Picture         =   "Critnote.frx":3B74
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   90
            Width           =   1215
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Enabled         =   0   'False
            Height          =   720
            Left            =   4485
            Picture         =   "Critnote.frx":4758
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   90
            Width           =   1095
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   720
            Left            =   3375
            Picture         =   "Critnote.frx":533C
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   90
            Width           =   1095
         End
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Height          =   720
            Left            =   2265
            Picture         =   "Critnote.frx":58C6
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   108
            Width           =   1095
         End
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&dit"
            Height          =   720
            Left            =   1140
            Picture         =   "Critnote.frx":64AA
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   90
            Width           =   1095
         End
         Begin VB.CommandButton Commandhelp 
            Caption         =   "Help"
            Height          =   435
            Left            =   -1080
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   800
         End
      End
      Begin VB.ComboBox customercode 
         Height          =   1488
         Left            =   9300
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Top             =   864
         Visible         =   0   'False
         Width           =   4452
      End
      Begin VB.ComboBox Genledger 
         Height          =   315
         Left            =   5250
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   7140
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5076
         Left            =   312
         TabIndex        =   15
         Top             =   2040
         Width           =   13380
         _ExtentX        =   23601
         _ExtentY        =   8954
         _Version        =   393216
         BackColorFixed  =   12058623
         BackColorBkg    =   16777215
         FillStyle       =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox weight 
         Height          =   285
         Left            =   9240
         TabIndex        =   12
         Top             =   1530
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox freight 
         Height          =   285
         Left            =   5535
         TabIndex        =   9
         Top             =   1530
         Width           =   1095
         _ExtentX        =   1947
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox biltno 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   870
         Width           =   1905
         _ExtentX        =   3366
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox bundles 
         Height          =   285
         Left            =   7755
         TabIndex        =   11
         Top             =   1530
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   487
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
         Left            =   1440
         TabIndex        =   1
         Top             =   870
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
         Left            =   765
         TabIndex        =   34
         Top             =   2580
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1566
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rate 
         Height          =   285
         Left            =   765
         TabIndex        =   35
         Top             =   4110
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3260
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox amount 
         Height          =   285
         Left            =   735
         TabIndex        =   36
         Top             =   4440
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3260
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_NO 
         Height          =   315
         Left            =   315
         TabIndex        =   0
         Top             =   840
         Width           =   1110
         _ExtentX        =   1969
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox marka 
         Height          =   285
         Left            =   6645
         TabIndex        =   10
         Top             =   1530
         Width           =   1065
         _ExtentX        =   1884
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox textbox 
         Height          =   312
         Left            =   9300
         TabIndex        =   5
         Top             =   840
         Width           =   4464
         _ExtentX        =   7874
         _ExtentY        =   550
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rsno 
         Height          =   285
         Left            =   11625
         TabIndex        =   14
         Top             =   1515
         Width           =   1695
         _ExtentX        =   2985
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox bdated 
         Height          =   315
         Left            =   4470
         TabIndex        =   3
         Top             =   870
         Width           =   1050
         _ExtentX        =   1842
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtschool 
         Height          =   315
         Left            =   2340
         TabIndex        =   8
         Top             =   1500
         Width           =   3165
         _ExtentX        =   5588
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin VB.ComboBox txtMark 
         Height          =   315
         ItemData        =   "Critnote.frx":68EC
         Left            =   10830
         List            =   "Critnote.frx":68F9
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1515
         Width           =   810
      End
      Begin MSMask.MaskEdBox txtTODDate 
         Height          =   315
         Left            =   2220
         TabIndex        =   67
         Top             =   8460
         Width           =   1035
         _ExtentX        =   1820
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNSCHNo 
         Height          =   315
         Left            =   5535
         TabIndex        =   4
         Top             =   870
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777152
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Checked :"
         Height          =   252
         Left            =   3852
         TabIndex        =   78
         Top             =   8496
         Width           =   684
      End
      Begin VB.Label lblfrt 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Freight :"
         Height          =   255
         Left            =   9720
         TabIndex        =   72
         Top             =   9360
         Width           =   660
      End
      Begin VB.Label lblPartyfrt 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10395
         TabIndex        =   71
         Top             =   9360
         Width           =   525
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NS-CH.No : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5625
         TabIndex        =   69
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label25 
         Caption         =   "V. No :"
         Height          =   255
         Left            =   420
         TabIndex        =   68
         Top             =   8460
         Width           =   795
      End
      Begin VB.Label Label24 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   420
         TabIndex        =   65
         Top             =   8100
         Width           =   1035
      End
      Begin VB.Label Label23 
         Caption         =   "Amt in words :"
         Height          =   255
         Left            =   9720
         TabIndex        =   63
         Top             =   8340
         Width           =   1035
      End
      Begin VB.Label lblMail 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9720
         TabIndex        =   61
         Top             =   9090
         Width           =   3855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "School "
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2220
         TabIndex        =   60
         Top             =   1260
         Width           =   780
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   924
         Left            =   312
         Top             =   8832
         Width           =   9312
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Godown "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10830
         TabIndex        =   58
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "R.S. No."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   11880
         TabIndex        =   57
         Top             =   1260
         Width           =   1200
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Agent :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   56
         Top             =   1260
         Width           =   510
      End
      Begin VB.Label labelbybanklbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "By Bank : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   55
         Top             =   7725
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label labelbybank 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4020
         TabIndex        =   54
         Top             =   7725
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4365
         TabIndex        =   53
         Top             =   570
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dem."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6645
         TabIndex        =   52
         Top             =   1290
         Width           =   1065
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Total Quantity : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4020
         TabIndex        =   51
         Top             =   7140
         Width           =   1155
      End
      Begin VB.Label tqu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4035
         TabIndex        =   50
         Top             =   7485
         Width           =   1140
      End
      Begin VB.Label mgd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11430
         TabIndex        =   49
         Top             =   7680
         Width           =   1200
      End
      Begin VB.Label mna 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11430
         TabIndex        =   48
         Top             =   7950
         Width           =   1200
      End
      Begin VB.Label mga 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10140
         TabIndex        =   47
         Top             =   7680
         Width           =   1170
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   46
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Note No. : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   315
         TabIndex        =   45
         Top             =   555
         Width           =   1155
      End
      Begin VB.Label label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9300
         TabIndex        =   44
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Net Amount : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10155
         TabIndex        =   43
         Top             =   7950
         Width           =   1260
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Gross Amount : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10155
         TabIndex        =   42
         Top             =   7380
         Width           =   1260
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bundle(s) : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7755
         TabIndex        =   41
         Top             =   1290
         Width           =   1515
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bilty No. "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2565
         TabIndex        =   40
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Freight : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5535
         TabIndex        =   39
         Top             =   1290
         Width           =   1155
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Weight : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9345
         TabIndex        =   38
         Top             =   1290
         Width           =   1560
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Total Discount : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11385
         TabIndex        =   37
         Top             =   7380
         Width           =   1290
      End
   End
End
Attribute VB_Name = "Critnote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kk As ADODB.Recordset
Dim I As Integer
Dim lastrow, lastcol As Integer
Dim VALIDRATE As Boolean
Dim maxrow As Integer
Public totalamount, totaldiscount As Double
Public otheramount, otherdiscount As Double
Dim autoscroll As Boolean
Public Edit As Boolean
Dim addmode As Boolean
Dim addoredit As Boolean
Dim Printheader As Boolean
Dim emptyInv_bool As Boolean
Dim entryNo_ As String
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
Sub printinvoice()

Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = False
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = False
Me.Commandabandon.Enabled = True
Me.CommandPrint.Enabled = True
Me.Command1.Enabled = True
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
Dim LEFTM As Integer
Open "" + VB.App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
FooterYes = False
header:
If kkk.State = 1 Then kkk.close
    CNSetup
    kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
    If FooterYes = True Then
        If Line > MaxLine - 5 Then
            Do While Line < 61
                Print #1, ""
                Line = Line + 1
            Loop
        End If
        FooterYes = False
        Line = 0
        LEFTM = 5
        Print #1, Tab(0); repli("-", 96)
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
Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CREDIT NOTE")))) / 2 - 3); Chr(14); "CREDIT NOTE"; Chr(20); Tab(54); IIf(Printheader = True, kkk!uptt, "")
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
If rs1.State = 1 Then rs1.close
rs1.Open "CREDITA", con, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.text) + "'"
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); " To,  S.L. Code : "; Tab(20); Mid$(rs1!subledger, 1, 5); Tab(46); "C/Note No. : "; Trim(rs1!invoiceNo); Tab(74); "Dated     : "; Chr(27) + Chr(72); rs1!invoiceDate
    If kkk.State = 1 Then kkk.close
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
       Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE; Tab(45); Chr(27) + Chr(71); "Bilty No.  : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(75); Chr(27) + Chr(71); "Dated     : "; Chr(27) + Chr(72); IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
       Print #1, Tab(3); IIf(IsNull(kkk!address1), " ", kkk!address1); Tab(45); Chr(27) + Chr(71); "Freight    : "; Chr(27) + Chr(72); Trim(rs1!freight); Tab(75); Chr(27) + Chr(71); "Bundle(s) : "; Chr(27) + Chr(72); Trim(rs1!bundles)
       Print #1, Tab(3); IIf(IsNull(kkk!address2), " ", kkk!address2); Tab(45); Chr(27) + Chr(71); "Demurrage  : "; Chr(27) + Chr(72); Trim(rs1!marka); Tab(75); Chr(27) + Chr(71); "Weight    : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."
       Print #1, Tab(3); IIf(IsNull(kkk!address3), " ", kkk!address3); Chr(27) + Chr(72); Tab(45); Chr(27) + Chr(71); "Agent Name : "; Chr(27) + Chr(72); Trim(rs1!agentname); Tab(75); Chr(27) + Chr(71); " R.S. No.  : "; Chr(27) + Chr(72); Trim(rs1!rsno)
       kkk.close
       Print #1, Chr(27) + Chr(71); repli("-", 96)
       Print #1, Tab(0); "S.No."; Tab(15); "Book Description"; Tab(50); "Quantity"; Tab(62); "Rate"; Tab(74); "Amount"; Tab(86); "Net Amount"
       Print #1, repli("-", 96); Chr(27) + Chr(72)
       Line = Line + 8
    End If
    If called1 Then
        called1 = False
        GoTo printagain1
    End If
    If called2 Then
        called2 = False
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.close
    kk.Open "select * from CREDITB where invoiceno=" + Trim(rs1!invoiceNo) + " and " & stringyear & " order by printorder,sno ", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
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
               tdata.close
               If Not kk.EOF Then
                  sno = sno + 1
                  kk.MoveNext
               End If
               If kk.EOF Then
                    Exit Do
               End If
            Loop
            If Line > MaxLine - 5 Then
                    called2 = True
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
printagain2:
                    
                    called2 = False
                End If
                Print #1, Tab(70); repli("-", 12)
                Line = Line + 1
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CREDITB where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and printorder =" + Trim(Str(cdiscount)) + " group by printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(Str(tdata(0)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(vdis), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(70); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.close
             Loop
           End If
       End If
       Print #1, repli("-", 96)
       Line = Line + 1
       Print #1, Tab(50); rsets(Trim(Str(totalquantity)), 7); Tab(84); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
       Line = Line + 1
       If kk.State = 1 Then kk.close
       kk.Open "Select * from CreditC where " & stringyear & " and  invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
       If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount - kk!amount
                    Else
                        netamount = netamount + kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(60); Trim(kk!text) + "    " + Trim(Format(Str(Round(kk!rate, 2)), "0.00")); Tab(84); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(60); Trim(kk!text); Tab(84); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
        End If
        Print #1, Tab(84); repli("-", 12)
        Print #1, Chr(27) + Chr(71); Tab(45); "NET AMOUNT Cr. TO YOUR A/C :"; Tab(85); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
        Line = Line + 2
        VNetamt = netamount
        If kk.State = 1 Then kk.close
        kk.Open "Select * from CreditA where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(60); kk!txt1; Tab(84); rsets(Trim(Format(Str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + Round(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(60); kk!txt2; Tab(84); rsets(Trim(Format(Str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + Round(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(60); "BY BANK "; Tab(84); rsets(Trim(Format(Str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - Round(kk!baa, 2)
             End If
        End If
        Print #1, Tab(84); repli("-", 12)
        Line = Line + 1
      ' PRINT THE FOOTER IN INVOICE START
        Do While Line < 61
            Print #1, ""
            Line = Line + 1
        Loop
        Print #1, Tab(0); Chr(27) + Chr(71); toword(Round(VNetamt, 2)); Chr(27) + Chr(72)
        Print #1, Tab(0); repli("-", 96)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        CNSetup
        tempdata.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(65); "FOR " + Trim(tempdata!cname)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Close #1
        PrintOption.Show
End Sub

Sub princrnotecon()
Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = False
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = False
Me.Commandabandon.Enabled = True
Me.CommandPrint.Enabled = True
Me.Command1.Enabled = True
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
Dim LEFTM As Integer
Open "" + VB.App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
FooterYes = False
header:
If kkk.State = 1 Then kkk.close
    CNSetup
    kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
    If FooterYes = True Then
        If Line > MaxLine - 5 Then
            Do While Line < 61
                Print #1, ""
                Line = Line + 1
            Loop
        End If
        FooterYes = False
        Line = 0
        LEFTM = 5
        Print #1, Tab(0); repli("-", 96)
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); kkk!COURT; Tab(65); "FOR " + Trim(kkk!cname)
        Print #1, ""
        Print #1, Tab(1); Chr(27) + Chr(71); "Continued on Page : " & Pno; Chr(27) + Chr(72)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
       ' Print #1, ""
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
Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CREDIT NOTE")))) / 2 - 3); Chr(14); "CREDIT NOTE"; Chr(20); Tab(54); IIf(Printheader = True, kkk!uptt, "")
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
If rs1.State = 1 Then rs1.close
rs1.Open "CREDITA", con, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.text) + "'"
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); " To,  S.L. Code : "; Tab(20); Mid$(rs1!subledger, 1, 5); Tab(46); "C/Note No. : "; Trim(rs1!invoiceNo); Tab(74); "Dated     : "; Chr(27) + Chr(72); rs1!invoiceDate
    If kkk.State = 1 Then kkk.close
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
       Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE; Tab(45); Chr(27) + Chr(71); "Bilty No.  : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(75); Chr(27) + Chr(71); "Dated     : "; Chr(27) + Chr(72); IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
       Print #1, Tab(3); IIf(IsNull(kkk!address1), " ", kkk!address1); Tab(45); Chr(27) + Chr(71); "Freight    : "; Chr(27) + Chr(72); Trim(rs1!freight); Tab(75); Chr(27) + Chr(71); "Bundle(s) : "; Chr(27) + Chr(72); Trim(rs1!bundles)
       Print #1, Tab(3); IIf(IsNull(kkk!address2), " ", kkk!address2); Tab(45); Chr(27) + Chr(71); "Demurrage  : "; Chr(27) + Chr(72); Trim(rs1!marka); Tab(75); Chr(27) + Chr(71); "Weight    : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."
       Print #1, Tab(3); IIf(IsNull(kkk!address3), " ", kkk!address3); Chr(27) + Chr(72); Tab(46); Chr(27) + Chr(71); "Agent Name : "; Chr(27) + Chr(72); Trim(rs1!agentname); Tab(50); Chr(27) + Chr(71); "(" & txtMark & ")"; Tab(76); " R.S. No.  : "; Chr(27) + Chr(72); Trim(rs1!rsno)
       kkk.close
       Print #1, Chr(27) + Chr(71); repli("-", 96)
       Print #1, Tab(0); "S.No."; Tab(15); "Book Description"; Tab(50); "Quantity"; Tab(62); "Rate"; Tab(74); "Amount"; Tab(86); "Net Amount"
       Print #1, repli("-", 96); Chr(27) + Chr(72)
       Line = Line + 8
    End If
    If called1 Then
        called1 = False
        GoTo printagain1
    End If
    If called2 Then
        called2 = False
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.close
    kk.Open "SELECT invoiceno, bookcode, rate, sum(quantity) AS qty, sum(amount) as amt,printorder, discount,sno From CREDITB Where " & stringyear & " and invoiceno = " + Trim(rs1!invoiceNo) + " GROUP BY bookcode, rate, invoiceno, printorder, discount,sno ORDER BY printorder,sno", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where " & stringyear & "  and bookcode='" + Trim(kk!Bookcode) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
                Print #1, Tab(0); rsets(Trim(Str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(Str(kk!qty)), 5); Tab(58); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(Str(kk!amt), "0.00")), 12)
                totalquantity = totalquantity + kk!qty
                Line = Line + 1
                If Line > MaxLine - 3 Then
                    called1 = True
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
printagain1:
                    called1 = False
               End If
               tdata.close
               If Not kk.EOF Then
                  sno = sno + 1
                  kk.MoveNext
               End If
               If kk.EOF Then
                    Exit Do
               End If
            Loop
            If Line > MaxLine - 5 Then
                    called2 = True
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
printagain2:
                    
                    called2 = False
                End If
                Print #1, Tab(70); repli("-", 12)
                Line = Line + 1
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CREDITB where invoiceno=" + Trim(rs1!invoiceNo) + " and printorder =" + Trim(Str(cdiscount)) + " and " & stringyear & " group by printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(Str(tdata(0)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(vdis), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(70); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.close
             Loop
           End If
       End If
       Print #1, repli("-", 96)
       Line = Line + 1
       Print #1, Tab(50); rsets(Trim(Str(totalquantity)), 7); Tab(84); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
       Line = Line + 1
       If kk.State = 1 Then kk.close
       kk.Open "Select * from CreditC where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
       If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount - kk!amount
                    Else
                        netamount = netamount + kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(60); Trim(kk!text) + "    " + Trim(Format(Str(Round(kk!rate, 2)), "0.00")); Tab(84); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(60); Trim(kk!text); Tab(84); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
        End If
        Print #1, Tab(84); repli("-", 12)
        'Print #1, Chr(27) + Chr(71); Tab(45); "NET AMOUNT Cr. TO YOUR A/C :"; Tab(85); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
        Print #1, Chr(27) + Chr(71); Tab(45); "NET AMOUNT Cr. TO YOUR A/C :"; Tab(85); rsets(Trim(Format(Str(Round(mna, 2)), "0.00")), 12); Chr(27) + Chr(72)
        
        Line = Line + 2
        VNetamt = netamount
        If kk.State = 1 Then kk.close
        kk.Open "Select * from CreditA where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(60); kk!txt1; Tab(84); rsets(Trim(Format(Str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + Round(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(60); kk!txt2; Tab(84); rsets(Trim(Format(Str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + Round(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(60); "BY BANK "; Tab(84); rsets(Trim(Format(Str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - Round(kk!baa, 2)
             End If
        End If
        Print #1, Tab(84); repli("-", 12)
        Line = Line + 1
      ' PRINT THE FOOTER IN INVOICE START
        Do While Line < 61
            Print #1, ""
            Line = Line + 1
        Loop
        Print #1, Tab(0); Chr(27) + Chr(71); toword(Round(VNetamt, 2)); Chr(27) + Chr(72)
        Print #1, Tab(0); repli("-", 96)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        CNSetup
        tempdata.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(65); "FOR " + Trim(tempdata!cname)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Close #1
        PrintOption.Show

End Sub
  
Sub CREDITCalc()
    'OTHERcredit.calc
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     mna.Caption = Format(Round((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
End Sub
Sub CREDITAbandon()
On Error Resume Next

txtchecked.text = ""

        Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
        Command1.Enabled = True
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
If kk.State = 1 Then
   kk.close
End If
   If Edit = False Then
   
    
    
   
    End If
        Dim ctl As Control
        For Each ctl In Me.Controls
            If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
                If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) Then
                    ctl.text = ""
                End If
                ctl.Enabled = False
                Critnote.customercode.Enabled = False
            End If
        Next
        
        For I = 1 To maxrow
           Grid1.Row = I
            For J = 1 To 8
                Grid1.Col = J
               Grid1.text = ""
           Next
        Next
        I_DTOB = "__/__/____"
        bdated = "__/__/____"
        tqu.Caption = ""
        mga.Caption = ""
        mgd.Caption = ""
        mna.Caption = ""
        labelbybank.Caption = ""
        maxrow = 0
        addoredit = False
        lblPartyfrt.Caption = ""
        txtTODNO.text = ""
        txtTODDate.text = "__/__/____"
        
        txtTODNO.Enabled = False
        txtTODDate.Enabled = False
        
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
                
                If InStr(Me.textbox.text, "(EM)") > 0 Then
                   Set RS = con.Execute("exec BookSearch_bycode_EM '" & session & "'," & main.setupid & ",'" & Trim(Grid1.text) & "'")
                Else
                   Set RS = con.Execute("exec BookSearch_bycode_NonEM '" & session & "'," & main.setupid & ",'" & Trim(Grid1.text) & "'")
                End If
                
                Row = Grid1.Row
                Col = Grid1.Col
                If Trim(Grid1.text) <> "" Then
                     
                  'If InStr(Me.textbox.Text, "(EM)") > 0 Then
                  
                   If (Trim(Grid1.text) <> "") Then
                     If RS.BOF = True Then
                        MsgBox "Book is note valid for this Party ...", vbCritical
                        Exit Function
                     End If
                    End If
                  
                  'End If
                                  
                  If Not RS.BOF Then
                  
                  
                        'RS.MoveFirst
                        'RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF Then
                            tempmeb.Visible = True
                            tempmeb.SetFocus
                            RS.close
                            templost = False
                            Exit Function
                        Else
                            Grid1.text = RS(0)
                            Grid1.Col = 2
                            Grid1.text = RS(1)
                            
                         'If Not edit Then
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
                                If Grid1.text = "" And addmode = True Then
                                     If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        kk.close
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
                                     
                                     
                                     
                                     
                                '===============================================
                                    'New Code For Sub Group
                                    If IsEmpty(tempstr) Then tempstr = ""
                                    s_ = tempstr
                                    
                                    If rs1.State = 1 Then rs1.close
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
                                 
                                 '===================Serise Wise============================
                                 D = ReturnDiscountNew_Return(RS(0), Trim(customercode.text), txtScId.text)
                                     If D > 0 Then
                                        Grid1.Col = 4
                                        Grid1.text = Format(D, "    0.00")
                                        Grid1.Col = 6
                                        Grid1.text = Format(D, "0.00")
                                        r = RS(3)
                                    End If
                                 '==========================================================
                                 

                                     
                                     
                                     
                                    Grid1.Col = 7
                                    Grid1.text = Format(Round(q * r, 2), "0.00")
                                    Grid1.Col = 8
                                    Grid1.text = Format(Round((q * r) * (D / 100), 2), "0.00")
                            Else
                               If Grid1.text = "" And addmode = False Then
                                     If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        kk.close
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
                                        
                                        
                                        
                                        
                                    '===============================================
                                    'New Code For Sub Group
                                    If IsEmpty(tempstr) Then tempstr = ""
                                    s_ = tempstr
                                    
                                    If rs1.State = 1 Then rs1.close
                                    rs1.Open "select GROUPCODE_sub from books where bookcode='" & RS(0) & "'", con
                                    If rs1.EOF = False Then
                                    If (Not IsNull(rs1!GROUPCODE_sub) And Len(rs1!GROUPCODE_sub) > 0) Then
                                        D = ReturnDiscount("" & category, "" & s_, Trim(rs1(0)))
                                        If D > 0 Then
                                            Grid1.Col = 4
                                            Grid1.text = Format(D, "    0.00")
                                            Grid1.Col = 6
                                            Grid1.text = Format(D, "0.00")
                                            D = D
                                            r = RS(3)
                                        End If
                                    End If
                                    End If
                                  'End Code For Sub Group
                                 '===============================================

                                 '===================Serise Wise============================
                                 D = ReturnDiscountNew_Return(RS(0), Trim(customercode.text), txtScId.text)
                                     If D > 0 Then
                                        Grid1.Col = 4
                                        Grid1.text = Format(D, "    0.00")
                                        Grid1.Col = 6
                                        Grid1.text = Format(D, "0.00")
                                        r = RS(3)
                                    End If
                                 '==========================================================


                                        
                                        
                                        
                            
                                End If
                            
                                End If
                            
                            End If
                            Grid1.Col = Col
                            RS.close
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
        
        
        totalamount = 0
        totaldiscount = 0
        
        For I = 1 To maxrow
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Val(Trim(Grid1.text))
            Grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.text))
        Next
        CREDITCalc
        Me.tqu.Caption = ""
        For I = 1 To maxrow
            Grid1.Col = 3
            Grid1.Row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.text))
        Next
        
        
        
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
                           If Commandother.Enabled = True Then
                               Commandother.SetFocus
                           End If
                        Exit Sub
                    End If
                End If
                Grid1.Row = Row
                Grid1.Col = Col
                Grid1.text = Bookname.text
                '/*************************
                If RS.State = 1 Then
                    RS.close
                End If
                
                
                
                RS.Open "books", con, adOpenDynamic, adLockReadOnly, adCmdTable
                Row = Grid1.Row
                Col = Grid1.Col
                If Trim(Grid1.text) <> "" Then
                    If Not RS.BOF Then
                        RS.MoveFirst
                        RS.Find "bookname='" + Trim(Grid1.text) + "'"
                        If RS.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            RS.close
                            Exit Sub
                        Else
                            
                            Grid1.Col = 1
                            Grid1.text = RS(0)
                            Grid1.Col = 2
                            Grid1.text = RS(1)
                        '    If Not edit Then
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
                                    kk.close
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
                                
                                
                                
                                    '===============================================
                                    'New Code For Sub Group
                                    If IsEmpty(tempstr) Then tempstr = ""
                                    s_ = tempstr
                                    
                                    If rs1.State = 1 Then rs1.close
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
                            RS.close
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
            totalamount = totalamount + Val(Trim(Grid1.text))
            Grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.text))
        Next
        CREDITCalc
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


Private Sub cmbAgentName_GotFocus()
'cmbAgentName.ListIndex = 0
End Sub

Private Sub cmbAgentName_LostFocus()
If cmbAgentName.text = "" Then
   MsgBox "Enter a Agent Name.. "
   cmbAgentName.text = "."
   'cmbAgentName.SetFocus
   Exit Sub
Else
''  Dim rs1 As New ADODB.Recordset
''  rs1.Open "select *  from AgentMaster where AgentName='" & cmbAgentName.Text & "' and " & stringyear & " order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
''  If rs1.RecordCount <= 0 Then
''     MsgBox "Enter valid Agent Name.. "
''     cmbAgentName.SetFocus
''  End If
  Dim rs1 As New ADODB.Recordset
  rs1.Open "select rep  from SalesRepQry where rep='" & cmbAgentName.text & "'", CON_blue
  If rs1.EOF = True Then
     MsgBox "Enter valid Agent Name.. "
     cmbAgentName.SetFocus
  End If
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
Set RS = con.Execute("exec searchList 'CNAndCI'")

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

Private Sub Command1_Click()
printch = "CREDITA"
ino = I_NO
printch1 = "INVOICENO"


Printheader = False
princrnotecon
End Sub

Private Sub Command2_Click()


    printch = "CREDITA"
    ino = I_NO
    printch1 = "INVOICENO"
    
    Printheader = True
    If Check1_dos.value = 1 Then
       printButton = "2"
       princrnotecon
    Else
       printButton = "1"
       PrintOption.Show
    End If




End Sub

Private Sub Commandabandon_Click()
CREDITAbandon

mnuMenu_ = "mnuCreditNItem"
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

End Sub
Private Sub Commandadd_Click()
On Error Resume Next
    
    entryNo_ = txtNSCHNo.text
    
    addmode = True
    Edit = False
    CREDITAbandon
    Dim RS As ADODB.Recordset
    addoredit = True
    addmode = True
    Set RS = New ADODB.Recordset
    Dim TEMPNUM As Integer
    If Edit = False Then
       'If CON.Execute("Select max(invoiceno) from CREDITA")(0) >= Val(Trim(Me.I_NO.Text)) Then
              Me.I_NO.text = con.Execute("Select max(invoiceno) from CREDITA")(0) + 1
              RS.Open "tempCRITNOTE", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
              If RS.BOF Then
                 RS.AddNew
              End If
               Me.I_NO.text = RS!In + 1
               RS!In = Val(Me.I_NO.text)
               RS.update
               RS.close
        'End If
    End If
    Dim ctl As Control
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = True
        End If
        'If UCase(Trim(ctl.Name)) = UCase(Trim(Me.I_NO.Name)) Then
        '   ctl.Enabled = False
        'End If
    Next
    Picture5.Enabled = True
    Commandother.Enabled = True
    Commandadd.Enabled = False
    Commanddelete.Enabled = False
    Commandedit.Enabled = False
    CommandPrint.Enabled = False
    Command1.Enabled = False
    Commandall.Enabled = True
    
    Commandsearch.Enabled = False
    Commandsave.Enabled = False
    Grid1.Enabled = True
    Commandsave.Enabled = False
    Me.customercode.Enabled = True
    
    txtMark.ListIndex = 0
    txtNSCHNo.text = entryNo_
    'i_dt.SetFocus
    I_NO.SetFocus
    
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
End Sub
Function returnCategory(s As String) As String
    Dim s1 As New ADODB.Recordset
    If s1.State = 1 Then s1.close
    
    s1.Open "select category from [groups] where groupcode='" & s & "' and " & stringyear, con
    If s1.EOF = False Then
       returnCategory = s1(0)
    End If
    
End Function

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
        RS.close
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
            Grid1.text = Format(RS(3), "0.00")            'rs(3)
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

            
            'Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
            Grid1.Col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.close
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
        '    Grid1.col = col
    End If
    RS.close
   ' row = Grid1.row
   ' col = Grid1.col
    totalamount = 0
    totaldiscount = 0
    Me.tqu.Caption = ""
    For I = 1 To Grid1.rows - 1
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Val(Trim(Grid1.text))
            Grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.text))
            Grid1.Col = 3
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.text))
     Next
     maxrow = Grid1.rows - 1
Else
'Grid1_Click
Exit Sub
End If

CREDITCalc

End Sub

Private Sub Commanddelete_Click()

On Error GoTo Del

Dim rs_h As New ADODB.Recordset
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

If rs1.State = 1 Then rs1.close
rs1.Open "select * from credita where INVOICENO=" & Trim(I_NO.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
If rs1.EOF = False Then
   
If rs_h.State = 1 Then rs_h.close
rs_h.Open "select * from credita where INVOICENO=" & Trim(I_NO.text) & " and " & stringyear, con
'If rs_h.Fields("Print_yes").Value = "y" Then
   If rs1!bAuthorized = True Then
      MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
      Exit Sub
   End If

'End If

End If

createLog UserName, I_NO, "Credit Note Item ", " Delete : " & mna.Caption, Date
    
If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else


        If (AuditTrail = "y") Then
        
        If (txtchecked.text = "y") Then
        
            actionType_ = "Delete"
            vtype1_ = "CI"
            vtypeNew = "CI"
            vdate_ = Trim(i_dt.text)
            vno_ = Trim(I_NO.text)
            
            frmAuditTrailLog_Rem.Show 1
            
         End If
        
        End If

        con.Execute ("delete  from CREDITA where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
        con.Execute ("delete  from CREDITB where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
        con.Execute ("delete  from CREDITC where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
        con.Execute ("delete from CREDITB_Free where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
        
        CREDITAbandon
End If


Commanddelete.Enabled = False
Commandsave.Enabled = False

Commandadd.SetFocus


Exit Sub
Del:
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub Commandedit_Click()

    Commandadd.Enabled = False
    'Me.Commandedit.Enabled = False
    Picture5.Enabled = True
    Commandother.Enabled = True
    Commandadd.Enabled = False
    Commandedit.Enabled = False
    Commandall.Enabled = True
    Commandsave.Enabled = False
    Commanddelete.Enabled = False
    Commandsearch.Enabled = False
    CommandPrint.Enabled = False
    Command1.Enabled = False
    Grid1.Enabled = True
    Me.customercode.Enabled = True
    Edit = True
    I_NO_LostFocus
    i_dt.Enabled = True
    i_dt.SetFocus
    
   '' I_NO.SetFocus   ''vk
     
    ' CREDITCtmp creation start
    DoEvents
    con.Execute ("delete from CREDITCtmp where INVOICENO = " & Critnote.I_NO & " and " & stringyear)
    con.Execute ("insert into CREDITCtmp(INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate," & _
    "AMOUNT,DEBITORCREDIT,TEXT,RYN,fyear," & _
    "setupid)  select INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate," & _
    "AMOUNT,DEBITORCREDIT,TEXT,RYN,fyear," & _
    "setupid from CREDITC where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
    DoEvents
    ' CREDITTMP creation end
    addoredit = False
    HIT
    Dim kx As Integer
    kx = 0
    Do While kx < 18000
    kx = kx + 1
    Loop
    DoEvents
    
    searchForm = "credititem"
    
    panel.Enabled = True
    Me.Enabled = True
    
    
    Dim ctl As Control
    For Each ctl In Me.Controls
    
    If (TypeOf ctl Is Label Or TypeOf ctl Is textbox Or TypeOf ctl Is MaskEdBox Or TypeOf ctl Is ComboBox) Then
        ctl.Enabled = True
    End If
    
    Next
    
    Commandother.Enabled = True
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    Commandsave.Enabled = False
    'Commanddelete.Enabled = False
    
End Sub
Private Sub Commandother_Click()
    
mnuMenu_ = "mnuCreditNItem"
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
Commandsave.Enabled = True
    
searchForm = "credititem"
frmEndPartTrans.Show 1
   
End Sub
Private Sub CommandPrint_Click()

printch = "CREDITA"
ino = I_NO
printch1 = "INVOICENO"


Printheader = True
printinvoice

End Sub
Private Sub CommandReturn_Click()
   Dim RS As New ADODB.Recordset
   Commandsave.Enabled = True
   RS.Open "tempCRITNOTE", con, adOpenDynamic, adLockOptimistic, adCmdTable
   If RS.BOF Then
       RS.AddNew
   End If
   RS!In = con.Execute("Select max(invoiceno) from CREDITA")(0)
   RS.update
   RS.close
   
   Unload Me
   addoredit = False
   ''MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()
    
On Error GoTo save_


Dim Checked_YesNo  As Integer

 If (txtchecked.text = "y") Then
      Checked_YesNo = 1
 Else
      Checked_YesNo = 0
 End If


Dim rs_h As New ADODB.Recordset
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

If rs1.State = 1 Then rs1.close
rs1.Open "select top 100 * from credita where INVOICENO=" & Trim(I_NO.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
If rs1.EOF = False Then
If rs_h.State = 1 Then rs_h.close
rs_h.Open "select top 100 * from credita where INVOICENO=" & Trim(I_NO.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
   If rs1!bAuthorized = True Then
      MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
      Exit Sub
   End If

End If
    
'=====================

    
    
    
    
    Dim SAVED As Boolean
    Dim LAMOUNT As Double
    Dim RS As ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Set RS = New ADODB.Recordset
    
     If Edit = False And addmode = False Then
      Me.Commandsave.Enabled = False
      Exit Sub
    End If
    If MsgBox("Do you want to save it now ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    SAVED = False
    Grid1.Row = 1
    Grid1.Col = 1
    If Trim(Grid1.text) = "" Then
        MsgBox "Please Enter item.... "
        Exit Sub
    End If
        
    '----------------------------------------------------------------
    If Edit = False Then
       If check_Duplikate("CREDITA", I_NO.text) = True Then
          MsgBox "This Inv. Number Already Exist ..", vbCritical
          Exit Sub
       End If
    
    
       '''==========Check Credit Note=========
     Set rs3 = New ADODB.Recordset
     rs3.Open "Select CNN from cnf1a  where cnn = " & Val(I_NO.text) & " and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
     If rs3.RecordCount > 0 Then
                MsgBox "CREDIT Note File already exist..."
                I_NO.SetFocus
                HIT
                Exit Sub
     End If
    
    
    End If
    '----------------------------------------------------------------
    createLog UserName, I_NO, "Credit Note Item ", " Save : " & mna.Caption, Date
    
    
    If Trim(I_NO.text) <> "" And Trim(i_dt.text) <> "" And Trim(customercode.text) <> "" Then
        If Edit Then
            con.Execute ("delete from CREDITA where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
            con.Execute ("delete from CREDITB where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
            con.Execute ("delete from CREDITC where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
            con.Execute ("delete from CREDITB_Free where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
            
        End If
        If RS.State = 1 Then
            RS.close
        End If
            LAMOUNT = 0
            RS.Open "select * from CREDITA  where " & stringyear & " and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
            If Not Edit Then
again:
               If con.Execute("Select max(invoiceno) from CREDITA")(0) >= Val(Trim(Me.I_NO.text)) Then
                    'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'GoTo again
               End If
            End If
            RS.AddNew
            
            If (AuditTrail = "y") Then
            
            If (AuditTrail = "y") Then
            RS!Checked_YesNo = Checked_YesNo
            End If
            
            End If
            
            
            If txtTODNO.text <> "" Then
               RS!todid = txtTODNO.text
               RS!toddate = txtTODDate.text
            Else
               RS!todid = Null
               RS!toddate = Null

            End If
            
            RS!remarks = Trim(txtRem)
            RS!NsChallanNo = Trim(txtNSCHNo)
            
            
            RS!scname = Trim(txtschool.text)
            RS!scid = txtScId.text
            
            RS!invoiceNo = Val(Me.I_NO.text)
            
            RS!invoiceNo = Val(Me.I_NO.text)
            RS!Godown = txtMark.text
            RS!invoiceDate = Me.i_dt.text
            RS!Genledger = Trim(Me.Genledger.text)
            RS!subledger = Trim(Me.customercode.text)
            RS!marka = Trim(Me.marka.text)
            RS!bundles = Trim(Me.bundles)
            RS!biltyno = Trim(Me.biltno.text)
            If Trim(Me.bdated) = Trim("__/__/____") Then
                RS!BILTYDATE = Null '"__/__/____"
            Else
                RS!BILTYDATE = Me.bdated & ""
            End If

            RS!freight = Me.freight & ""
            RS!weight = Me.weight & ""
            RS!rsno = Me.rsno & ""
            RS!netamount = Round(Val(Trim(Me.mna.Caption)), 2)
            RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
            RS!txt1 = Trim(frmEndPartTrans.T1TEXT.text)
            RS!txt1a = Val(Trim(frmEndPartTrans.T1.text))
            RS!txt2 = Trim(frmEndPartTrans.T2TEXT.text)
            RS!txt2a = Val(Trim(frmEndPartTrans.T2.text))
            RS!baa = Val(Trim(frmEndPartTrans.T3TEXT.text))
            RS!agentname = cmbAgentName.text
            Dim trs As New ADODB.Recordset
            trs.Open " SELECT DISTCODE  FROM SLEDGER  WHERE SUBLEDGER='" & customercode.text & "'", con, adOpenStatic, adLockOptimistic, adCmdText
            If Not trs.BOF Then
               RS!District = trs!distcode & ""
            Else
               RS!District = ""
            End If
            trs.close
err1:
            If Not Edit Then
                If con.Execute("Select max(invoiceno) from CREDITA")(0) >= Val(Trim(Me.I_NO.text)) Then
                    'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'rs!INVOICENO = Val(Me.I_NO.Text)
                    On Error GoTo err1
                End If
            End If
            
            RS!Amtwords = Trim(txtAmtwords.text)
            RS!fyear = session
            RS!setupid = setupid
            RS.update
            
            '---------------------------------------------------------------------
                        
            
            
            On Error GoTo 0
            RS.close
            RS.Open "select * from CREDITB where " & stringyear & " and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
            Dim I As Integer
            RRRR = Grid1.Row
            CCCC = Grid1.Col
            For I = 1 To maxrow
                Grid1.Row = I
                Grid1.Col = 1
                If Trim(Grid1.text) <> "" Then
                    Grid1.Col = 3
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
                        RS!fyear = session
                        RS!setupid = setupid
                        
                        '============================================
                        If kk.State = 1 Then kk.close
                        kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,Apply from KitQry where kitcode='" & RS!Bookcode & "'", con
                        bkdesc = ""
                        While kk.EOF = False
                        If kk!Apply = "y" Then
                          con.Execute "insert into CREDITB_Free(INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
                          " values('" & Val(Me.I_NO.text) & "','" & Format(Me.i_dt.text, "MM/dd/yyyy") & "','" & Trim(Me.Genledger.text) & "','" & Trim(Me.customercode.text) & "','" & kk!Bookcode & "','" & (kk!qty * RS!QUANTITY) & "','" & kk!rate & "','" & Trim(Me.cmbAgentName.text) & "','" & setupid & "','" & session & "','" & txtMark & "')"
                          kk.MoveNext
                        End If
                        Wend
                        '==========================================
                        
                        
                        RS.update
                    End If
                End If
            Next
            RS.close
            Grid1.TopRow = 1
            
            RS.Open "select * from CREDITC where " & stringyear & " and invoiceno<=0", con, adOpenDynamic
            '/******
            Dim temprs As ADODB.Recordset
            Set temprs = New ADODB.Recordset
            For I = 1 To frmEndPartTrans.vs.rows - 1
                    frmEndPartTrans.vs.Row = I
                    frmEndPartTrans.vs.Col = 0
                    If Trim(frmEndPartTrans.vs.text) <> "" Then
                        RS.AddNew
                        RS!invoiceNo = Val(Me.I_NO.text)
                        RS!invoiceDate = Me.i_dt.text
                        RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
                        RS!text = Trim(frmEndPartTrans.vs.text)
                        If temprs.State = 1 Then
                            temprs.close
                        End If
                        If Edit Then
                        temprs.Open "select * from CREDITCtmp WHERE INVOICENO=" & Critnote.I_NO & "", con, adOpenStatic, adLockReadOnly, adCmdText
                        If frmEndPartTrans.vs.text <> "" Then
                                temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.text) + "'"
                                RS!Genledger = temprs!Genledger & ""
                                RS!subledger = temprs!subledger & ""
                                RS!DebitorCredit = temprs!DebitorCredit & ""
                                RS!RYN = temprs!RYN & ""
                        End If
                        temprs.close
                        Else
                        temprs.Open "select * from invoiceend where type='credititem' and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
                        If frmEndPartTrans.vs.text <> "" Then
                           temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.text) + "'"
                           RS!Genledger = temprs!Genledger & ""
                           RS!subledger = temprs!subledger & ""
                           RS!DebitorCredit = temprs!DebitorCredit & ""
                           RS!RYN = temprs!RYN & ""
                        End If
                        temprs.close
                        End If
                        frmEndPartTrans.vs.Col = 1
                        RS!rate = Val(Trim(frmEndPartTrans.vs.text))
                        If Val(Trim(frmEndPartTrans.vs.text)) > 0 Then
                            RS!amount = Round((Me.totalamount - Me.totaldiscount), 2) * Round((Val(Trim(frmEndPartTrans.vs.text)) / 100), 2)
                        Else
                        frmEndPartTrans.vs.Col = 2
                            RS!amount = Val(Trim(frmEndPartTrans.vs.text))
                        End If
                    RS!fyear = session
                    RS!setupid = setupid
    
                    RS.update
                    End If
            Next
            RS.close
            con.Execute ("delete  from CreditCTmp where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
            
         
         ss1_ = ""
         If rsno.text <> "" Then
            Set RS = New ADODB.Recordset
            RS.Open "select invoiceno from CREDITA where rsno='" & rsno & "' order by invoiceno", con
            While RS.EOF = False
               If ss1_ = "" Then
                  ss1_ = RS!invoiceNo
               Else
                  ss1_ = ss1_ & "," & RS!invoiceNo
               End If
               RS.MoveNext
           Wend
         End If
         
         If Len(rsno) > 0 Then
         con.Execute "update BILTYRETURNREGISTER set rno='" & ss1_ & "' where sno=" & rsno & ""
         End If
                    
         SAVED = True
           
        End If
        If SAVED Then
            MsgBox "Record Saved"
        Me.customercode.Enabled = False
        'Me.Grid1.Enabled = False
        Me.Commandall.Enabled = False
        Me.Commandother.Enabled = False
        Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = False
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
        Command1.Enabled = True
        End If
        addmode = False
        addoredit = False
        
        Me.Commandadd.SetFocus
        
        mnuMenu_ = "mnuCreditNItem"
        
        
        If (AuditTrail = "y") Then
        
        If (txtchecked.text = "y") Then
        
            actionType_ = "Edit"
            vtype1_ = "CI"
            vtypeNew = "CI"
            vdate_ = Trim(i_dt.text)
            vno_ = Trim(I_NO.text)
            
            frmAuditTrailLog_Rem.Show 1
            
         End If
        
        End If

        
        
        'Me.Commandsave.Enabled = False
        SetButton Commandadd, Commandedit, Commandsave, Commanddelete
        Commandsave.Enabled = False
        

Exit Sub
save_:
MsgBox "" & err.DESCRIPTION

        
End Sub
Private Sub Commandsave_GotFocus()
txtAmtwords = toword(Critnote.mna)
End Sub

Private Sub Commandsearch_Click()

sqlQry = "select InvoiceNo,InvoiceDate,Subledger,NetAmount from credita  where InvoiceNo"
orderby = "order by InvoiceNo"



searchType = "inv"
'popuplist10 "select InvoiceNo,InvoiceDate,Subledger,NetAmount from credita where " & stringyear & "  order by InvoiceNo", con
popuplistFast "select InvoiceNo,InvoiceDate,Subledger,NetAmount from InvoiceA where " & stringyear & "  order by InvoiceNo", con, , , "CI"

Check1_dos.Enabled = True
Check1_direct.Enabled = True
Check_header.Enabled = True
End Sub

Private Sub Commandsearch_GotFocus()

If PopUpValue1 <> "" Then
     I_NO.text = PopUpValue1
     I_NO_LostFocus
     PopUpValue1 = ""
     If I_NO.Enabled = True Then
     I_NO.SetFocus
     End If
     
     SetButton Commandadd, Commandedit, Commandsave, Commanddelete
     
End If

End Sub

Private Sub customercode_LostFocus()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
  
    
    RS.Open "select * from sledger where gledger='SUNDRY DEBTORS' and subledger='" + Trim(customercode.text) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If RS.RecordCount > 0 Then
       lblMail.Caption = RS!email & ""
       lblPartyfrt.Caption = RS!freight & ""
    End If
    
    If RS.State = 1 Then RS.close
    RS.Open "select top 100 * from sledger where subledger='" + Trim(customercode.text) + "' and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
    If RS.EOF Then
        customercode.SetFocus
        HIT
        RS.close
        Exit Sub
    End If
    
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    If RS!distcode <> "" And addmode = True Then
       rs1.Open "Select * from Districts where Districtname = '" & RS!distcode & "' and " & stringyear, con, adOpenStatic, adLockReadOnly
       If rs1.RecordCount > 0 Then
          Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
       End If
    End If
    If RS.State = 1 Then RS.close
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

mnuMenu_ = "mnuCreditNItem"
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
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
        Dim trs3 As New ADODB.Recordset
        trs3.Open "Select * from cnf1a  where cnn = " & Val(I_NO.text) & " and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
        If trs3.RecordCount > 0 Then
        End If
           If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("grid1")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("tempmeb")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bookname")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("rsno")) Then
              sendkeys ("{TAB}")
            End If
        End If
    End If
End Sub
Private Sub Form_Load()

Screen.MousePointer = vbHourglass

On Error Resume Next

Me.top = 0
Me.Left = 0

Me.Width = 13900
Me.Height = 10750

Me.Caption = "Credit Note Item"

i_dt.text = Format(Date, "dd/MM/yyyy")


Dim trs As ADODB.Recordset
Set trs = New ADODB.Recordset
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
   ' Set CON = New ADODB.Connection
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Me.top = 1470
    Me.Left = 90
    
    Me.top = 50
    Me.Left = 50
    
    Grid1.Left = 150
    
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
    Grid1.ColWidth(0) = 200
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 3800
    Grid1.ColWidth(3) = 1050
    Grid1.ColWidth(4) = 1050
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1500
    Grid1.ColWidth(8) = 1500
    Me.CommandPrint.Enabled = True
    
    
    
    
    If RS.State = 1 Then RS.close
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
    RS.close
    
    
    Genledger.text = "SUNDRY DEBTORS"
    'RS.Open "select * from sledger where gledger='" + Trim(Genledger.Text) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
    'Set RS = CON.Execute("exec fatch_ledger '" & Genledger.Text & "'")
    'Set RS = CON.Execute("exec fatch_ledger '" & Genledger.Text & "','" & session & "'," & main.setupid & "")
     If RS.State = 1 Then RS.close
     RS.Open "select * from sledger where gledger='" & Genledger.text & "' and " & stringyear, CCON, adOpenDynamic, adLockReadOnly, adCmdText

    
    If Not RS.BOF Then
        Do While Not RS.EOF
            Me.customercode.AddItem RS("subledger")
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close

      

''     '*******Agent  combo fill
''    RS.Open "select Distinct Agentname from AgentMaster  where " & stringyear & " order by agentname", CON, adOpenStatic, adLockReadOnly, adCmdText
''    cmbAgentName.Clear
''    If Not RS.EOF Then
''       Do While Not RS.EOF
''          If IsNull(RS(0)) = False Then
''            Me.cmbAgentName.AddItem RS(0)
''          End If
''          If Not RS.EOF Then RS.MoveNext
''        Loop
''    End If
''    RS.close

     '*******Agent  combo fill
    'popuplist10 "select Rep as Representative,Add1,Add2,District,[state] from SalesRepQry order by Rep", CON
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
    
 
 
 
    Dim rs_godwn As New ADODB.Recordset
    If rs_godwn.State = 1 Then rs_godwn.close
    rs_godwn.Open "select godwn from GodownMaster where Binder_Printer='g' order by id", con, adOpenDynamic, adLockReadOnly, adCmdText
    txtMark.Clear
    If Not rs_godwn.EOF Then
       Do While Not rs_godwn.EOF
          If IsNull(rs_godwn(0)) = False Then
            Me.txtMark.AddItem rs_godwn(0)
          End If
          If Not rs_godwn.EOF Then rs_godwn.MoveNext
        Loop
    End If
    
    txtMark.ListIndex = 0
  
    
    Bookname.Height = 1935
    Bookcode.Left = Grid1.Left
    Bookcode.Visible = False
    Bookname.Visible = False
    
    Grid1.rows = 200
    
    For I = 1 To 99
      Grid1.RowHeight(I) = 300
    Next
    
    Bookcode.Width = 1230
    Bookname.Width = 2830
    amount.Width = rate.Width
    
       
       kk.Open "SELECT MAX(INVOICENO) FROM CREDITA where " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
       If kk(0) <> "" Then
            Me.I_NO.text = kk(0)
            I_NO_LostFocus
       Else
            Me.I_NO.text = "1"
       End If
       kk.close
   
   'End If
   
   
   Commanddelete.Enabled = False
   Commandedit.Enabled = True
   Commandsave.Enabled = False
   Command1.Enabled = True
   lastrow = 0
   lastcol = 1
  
  Dim ctl As Control

   For Each ctl In Me.Controls
      If Not TypeOf ctl Is CommandButton Then
          ctl.Enabled = False
      End If
      If UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandall.Name)) Then
         ctl.Enabled = False
      End If
   Next
   
   
   
   Picture5.Enabled = True
   
   
   mnuMenu_ = "mnuCreditNItem"
'   SetButton Commandadd, Commandedit, Commandsave, Commanddelete
   
   Check1_dos.Enabled = True
   Check_header.Enabled = True
   Check1_direct.Enabled = True


Commandsave.Enabled = False
Commanddelete.Enabled = False

Screen.MousePointer = vbDefault
   
BackColorFrom Me

'Commandadd_Click

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Resize()

panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub

Private Sub Grid1_Click()

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
End Sub
Private Sub Grid1_KeyPress(KeyAscii As Integer)
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
                tempmeb.top = Grid1.top + Grid1.CellTop '- 50
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
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu dd, , Grid1.Left + X, Grid1.top + Y
End If
End Sub

Private Sub Grid1_Scroll()
If tempmeb.Visible <> True And Bookname.Visible <> True Then
        Bookname.Visible = False
        tempmeb.Visible = False
        Grid1.SetFocus
End If
autoscroll = True
End Sub

Private Sub i_dt_LostFocus()
On Error Resume Next
If Trim(i_dt.text) <> Trim("__/__/____") Then
    If Not checkdate(Trim(i_dt.text), i_dt) Then
       i_dt.Enabled = True
       i_dt.SetFocus
       Exit Sub
    End If
    Dim tRS1 As New ADODB.Recordset
    Dim trs2 As New ADODB.Recordset
    If trs2.State = 1 Then trs2.close
    trs2.Open "Select top 2 invoiceno as cn from credita where " & stringyear, con, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount <= 0 Then
       Exit Sub
    Else
        If tRS1.State = 1 Then tRS1.close
        tRS1.Open "Select top 10 min(invoiceno) as mid,invoicedate from credita where " & stringyear & " group by invoiceno,invoiceDate", con, adOpenDynamic, adLockOptimistic
        If tRS1.RecordCount > 0 Then
            If CDate(i_dt) <= tRS1!invoiceDate Then
               If CDate(i_dt) <> tRS1!invoiceDate Then
                If Month(CDate(i_dt)) <> 4 And Day(CDate(i_dt)) <> 1 Then
                 MsgBox "Please select valid Credit  No. for this date.."
                 i_dt.SetFocus
                 Exit Sub
                 Else
                 If tRS1!Mid <> 1 Then
                 If Val(I_NO) >= tRS1!Mid Then
                 MsgBox "Please select valid Invoice No. for this date.."
                 I_NO.SetFocus
                 Exit Sub
               End If
               End If
               End If
               End If
            End If
        End If
    End If
    If trs2.State = 1 Then trs2.close
    trs2.Open "Select top 10 max(invoiceno) as mid from credita where " & stringyear & " and convert(smalldatetime,invoicedate,103) <= convert(smalldatetime,'" & i_dt.text & "',103)-1", con, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount > 0 Then
        If IsNull(trs2!Mid) = False Then
            If Val(I_NO.text) >= trs2!Mid Then
               If tRS1.State = 1 Then tRS1.close
               tRS1.Open "Select top 10 min(InvoiceNo)as m2 from credita where " & stringyear & "  and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" & i_dt.text & "',103)+1", con, adOpenDynamic, adLockOptimistic
               If tRS1.RecordCount > 0 Then
                  If IsNull(tRS1!m2) <> True Then
                     If Val(I_NO.text) <= tRS1!m2 Then
                       
                     Else
                         MsgBox "Please select valid Credit No. for this date.."
                         I_NO.SetFocus
                     End If
                  End If
               End If
            
            Else
             If I_NO.Enabled = False Then Exit Sub
                    If i_dt.Enabled = True Then
                            MsgBox "Please select valid Invoice No for this date.."
                            I_NO.Enabled = True
                            I_NO.SetFocus
                    End If
                 End If
            End If
         
        
     End If
    Else
    i_dt.SetFocus
    HIT
    
    
    Exit Sub
End If
End Sub


Private Sub I_NO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'SendKeys "{tab}"
 
End If


If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Sub I_NO_LostFocus()

On Error Resume Next

Dim RS As ADODB.Recordset
Dim rs3  As ADODB.Recordset

Set RS = New ADODB.Recordset
Set rs3 = New ADODB.Recordset

If Val(inviceNo) > 0 Then
   I_NO.text = inviceNo
End If
inviceNo = ""



If Trim(I_NO.text) = "" Then
        MsgBox "Credit Note no Cannot be Null"
        I_NO.SetFocus
Else
    
    If RS.State = 1 Then RS.close
    RS.Open "Select * from  CREDITA where INVOICENO = " + Trim(I_NO.text) + " and " & stringyear, con, adOpenStatic, adLockReadOnly
    'RS.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
    
    rs3.Open "Select CNN from cnf1a  where cnn = " & Val(I_NO.text) & " and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
    If rs3.RecordCount > 0 Then
                MsgBox "CREDIT Note File already exist..."
                I_NO.SetFocus
                HIT
                Exit Sub
   End If
   
   
    If RS.EOF Then
     If addoredit = False Then
                'MsgBox "CREDIT NOT  not found"
                Exit Sub
            End If
            Exit Sub
    End If
   
   If addoredit Then
      MsgBox "CREDIT NOTE  already exist..."
      I_NO.SetFocus
      HIT
      Exit Sub
   End If
   
'   Dim ctl As Control
'   For Each ctl In Me.Controls
'       If Not TypeOf ctl Is CommandButton Then
'          ctl.Enabled = True
'       End If
'   Next
CREDITAbandon
        
        
        If (AuditTrail = "y") Then
        
            If (RS!Checked_YesNo = True) Then
               txtchecked.text = "y"
            Else
                txtchecked.text = "n"
            End If
        
        End If
        
        
        
        If Not IsNull(RS!todid) Or RS!todid = "" Then
           txtTODNO.text = RS!todid
           txtTODDate.text = RS!toddate
        End If

        
        Me.Commandother.Enabled = True
        
        txtschool.text = RS!scname & ""
        txtScId.text = RS!scid & ""
        
        txtRem = RS!remarks & ""
        txtNSCHNo = RS!NsChallanNo & ""
        
        txtAmtwords = RS!Amtwords & ""
        
        txtMark.text = RS!Godown & ""
        I_NO.text = RS!invoiceNo
        Me.i_dt.text = RS!invoiceDate
        Me.Genledger.text = Trim(RS!Genledger)
        Me.customercode.text = Trim(RS!subledger)
        Me.textbox.text = Trim(RS!subledger)
        Me.marka.text = Trim(IIf(IsNull(RS!marka), "", RS!marka))
        Me.bundles = Trim(RS!bundles)
        Me.biltno.text = Trim(RS!biltyno)
        
        If IsNull(RS!BILTYDATE) Then
           Me.bdated = "__/__/____"
        Else
           Me.bdated = RS!BILTYDATE
        End If
        Me.freight = Trim(RS!freight)
        Me.weight = Trim(RS!weight)
        Me.rsno = Trim(RS!rsno)
        Me.labelbybank = Trim(RS!baa)
        mna.Caption = RS!netamount
        Me.cmbAgentName.text = IIf(IsNull(RS!agentname), "", RS!agentname)
        RS.close
'*/**/*/*/*/*//*/*
        If RS.State = 1 Then
            RS.close
        End If
       
       RS.Open "Select top 1000 * from CREDITB where INVOICENO =" + Trim(I_NO.text) + " and " & stringyear & " order by SNO ", con, adOpenStatic, adLockReadOnly
        Grid1.TopRow = 1
        If Not RS.EOF Then
            Grid1.Row = 1
            Grid1.Col = 1
            Do While Not RS.EOF
               If Trim(RS!invoiceNo) = Trim(I_NO.text) Then
               Grid1.Col = 1
                Grid1.text = Trim(RS!Bookcode)
                If kk.State = 1 Then
                    kk.close
                End If
                kk.Open "select * from books where bookcode='" + Trim(RS!Bookcode) + "' and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
                If kk.EOF = True Then Exit Sub
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
        '    Me.i_dt.SetFocus
        End If
                Row = Grid1.Row
        Col = Grid1.Col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Val(Trim(Grid1.text))
            Grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.text))
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
    
    
    If I_NO <> "" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select mail,freight_yes_no from creditaQry where INVOICENO=" & I_NO & "", con
        If rs1.EOF = False Then
           lblMail.Caption = rs1!mail & ""
           lblPartyfrt.Caption = rs1!freight_yes_no & ""
        End If
    End If
    
    Me.Commandother.Enabled = False
    Me.Commanddelete.Enabled = False
    Me.Commandsave.Enabled = False



mnuMenu_ = "mnuCreditNItem"
'SetButton Commandadd, Commandedit, Commandsave, Commanddelete


End Sub

Private Sub I_OB_LostFocus()
I_OB = UCase(I_OB)
End Sub

Private Sub marka_LostFocus()
marka = UCase(marka)
End Sub

Private Sub station_LostFocus()
station = UCase(station)
End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub rsno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        fatchBiltyData
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
Sub fatchBiltyData()

If rsno.text <> "" Then
    If rs1.State = 1 Then rs1.close
    rs1.Open "select BDL,WT,Freight,Freight_Paid,gr,rr from BILTYRETURNREGISTER where SNO=" & rsno.text & "", con
    If rs1.EOF = False Then
    If RS.State = 1 Then RS.close
    RS.Open "select invoiceno from CREDITA where RSNO=" & rsno.text & "", con
    If RS.EOF = True Then
       bundles.text = rs1!BDL & ""
       freight.text = rs1!freight & ""
       weight.text = rs1!wt & ""
       marka.text = rs1!Freight_Paid & ""
       biltno.text = IIf(IsNull(rs1!gr), rs1!rr, rs1!gr)
    End If
    End If
End If
    
End Sub
Private Sub tempmeb_Change()
If Grid1.Col = 1 Or Grid1.Col = 2 Then
    Grid1.text = tempmeb.text
Else
    If Grid1.Col = 3 Then
        Grid1.text = Format(tempmeb.text, "0")
    Else
        Grid1.text = Format(tempmeb.text, "0.00")
    End If
End If
End Sub
Private Sub tempmeb_GotFocus()
    HIT
End Sub
Private Sub tempmeb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Dim RS As ADODB.Recordset
           Set RS = New ADODB.Recordset
            Select Case Grid1.Col
                Case 1
                    'If RS.State = 1 Then
                    '    RS.close
                    'End If
                    
                    If InStr(customercode.text, "(EM)") Then
                       Set RS = con.Execute("exec BookSearch_bycode_EM '" & session & "'," & main.setupid & ",'" & Trim(Grid1.text) & "'")
                    Else
                       Set RS = con.Execute("exec BookSearch_bycode_NonEM '" & session & "'," & main.setupid & ",'" & Trim(Grid1.text) & "'")
                       ''MsgBox "Non Em"
                    End If
                    
                    
                    If (Trim(Grid1.text) <> "") Then
                    
                        If RS.BOF = True Then
                           MsgBox "Book is note valid for this Party ...", vbCritical
                           Exit Sub
                        End If
                        
                    End If
                    
                    
                    'RS.Open "books", CON, adOpenStatic, adLockReadOnly, adCmdTable
                    If Not RS.BOF Then
                         
                         
                        'RS.MoveFirst
                        'RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF And Trim(Grid1.text) <> "" Then
                            RS.close
                            ''MsgBox "eof......"
                            
                            Exit Sub
                        Else
                            RS.close
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
                    
                    
                        'Grid1.Col = Grid1.Col + 2
                        Grid1.Col = 1
                        Grid1.Row = Grid1.Row + 1
                        Grid1.rows = Grid1.rows + 1
                        
                        
                        
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                Case 4
                    Grid1.Col = Grid1.Col + 2
              '       SendKeys "{LEFT}"
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
End Sub
Private Sub tempmeb_LostFocus()
    If templost Then
        tempmeb.Visible = False
    End If
End Sub
Private Sub textbox_GotFocus()
    Me.customercode.Enabled = True
    Me.customercode.Visible = True
  '' Me.customercode.Height = 1100
    Me.customercode.ZOrder
    Me.customercode.SetFocus
    
    '***********change  by vk
   ' If Me.Commandsave.Enabled = True Then
   '     Me.customercode.Enabled = True
    '    Me.customercode.Visible = True
        ' Me.customercode.Height = 1100
     '   Me.customercode.ZOrder
     '   Me.customercode.SetFocus
      
   ' End If
    
End Sub

Private Sub through_LostFocus()
    through = UCase(through)
End Sub

Private Sub through1_LostFocus()
through1 = UCase(through1)
End Sub
Sub fatchOrder()
      
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim dis
Dim party_ As String

party_ = ""



For I = 1 To Grid1.rows - 1
If Grid1.TextMatrix(I, 1) <> "" Then
   Grid1.Row = I
    For J = 1 To 8
        Grid1.Col = J
       Grid1.text = ""
   Next
End If
Next

'If PopUpValue1 = "" Then Exit Sub


Dim kk1 As Integer
Dim sqty As Integer
Dim oqty As Integer
Dim b1 As Boolean
b1 = False
sqty = 0
oqty = 0
kk1 = 1



Dim rate As Double
Dim qty As Double

rate = 0
dis = 0


If txtScId.text <> "" Then

'Set RS = con.Execute("exec Sp_FetchSaleReturn '" & last_dbase & "','" & txtScId.Text & "'")
Set RS = New ADODB.Recordset
'If RS.State = 1 Then RS.close
If txtNSCHNo.text <> "" Then

RS.Open "select a.scid,a.bcode,a.qty,a.billingDis,a.agentName,a.price,b.bookname from SchoolWiseBookReturnDet as a" & _
" inner join books as b on (a.BCode=b.BOOKCODE) where a.entryno=" & txtNSCHNo.text & " " & _
" and a.scid='" & txtScId.text & "' order by a.Id,b.bookname", con_LAST

While RS.EOF = False
    
    
If (txtScId.text = RS(0)) Then
    
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT QUANTITY FROM CreditbQry where (scid='" & txtScId.text & "' and BOOKCODE='" & RS!bcode & "' and rate=" & RS!Price & " and NsChallanNo=" & txtNSCHNo.text & ")", con
    If rs1.EOF = False Then
       qty = rs1(0)
    End If
    
    If (RS!qty - qty > 0) Then
    
        Grid1.TextMatrix(kk1, 1) = RS(1)     ' bcode
        'If rs3.State = 1 Then rs3.close
        'rs3.Open "select bookname,rate from books where bookcode='" & RS(1) & "'", con
        'If rs3.EOF = False Then
           Grid1.TextMatrix(kk1, 2) = RS!Bookname
           Grid1.TextMatrix(kk1, 5) = RS!Price
           rate = RS!Price
        'End If
        
        
        Grid1.TextMatrix(kk1, 3) = (RS!qty - qty)
        
        If Not IsNull(RS(3)) Then
           Grid1.TextMatrix(kk1, 4) = RS(3)
           Grid1.TextMatrix(kk1, 6) = RS(3)
        Else
           Grid1.TextMatrix(kk1, 4) = 0
           Grid1.TextMatrix(kk1, 4) = 0
        End If
    
    
    
        dis = RS(3)
        
        Grid1.TextMatrix(kk1, 7) = (Grid1.TextMatrix(kk1, 3) * rate)
        Grid1.TextMatrix(kk1, 8) = Format(Round(Grid1.TextMatrix(kk1, 7) * (dis / 100), 2), "0.00")
        
        kk1 = kk1 + 1

End If
    

End If

RS.MoveNext

Wend

End If
     
     

Dim totalamount_ As Double
Dim totaldiscount_ As Double

totalamount_ = 0
totaldiscount_ = 0
Qty_ = 0
qty = 0

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
      
End If
      
      
End Sub
Private Sub txtNSCHNo_GotFocus()

If PopUpValue1 <> "" Then

    'txtNSCHNo.Text = PopUpValue1
    txtScId = PopUpValue1
    
    txtschool.text = popupvalue5
    
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""
    popupvalue5 = ""
    

End If

End Sub

Private Sub txtNSCHNo_KeyDown(KeyCode As Integer, Shift As Integer)
 
 

' Set RS = con.Execute("exec Sp_FetchSaleReturn " & last_dbase & "")
 
 
'If (KeyCode = 113) Then
'popuplistFastNew "", con, , , "I"
'End If


If (KeyCode = 13) Then

    If txtNSCHNo.text = "" Then Exit Sub

    If RS.State = 1 Then RS.close
    RS.Open "select NoofGaddi,BiltyCharges,PartyName from schoolWiseBookReturn where entryno=" & txtNSCHNo.text, con_LAST
    If RS.EOF = False Then
       If Val(RS!BiltyCharges) > 0 Then
          freight.text = RS!BiltyCharges & ""
       End If
       
       'bundles.Text = RS!noofgaddi & ""
       
       textbox.text = RS!partyname & ""
       customercode.text = RS!partyname & ""
    End If
    
    
End If

 

End Sub

Private Sub txtRem_LostFocus()
txtRem = UCase(txtRem)
End Sub

Private Sub txtschool_GotFocus()
If RS.State = 1 Then RS.close
If PopUpValue1 <> "" Then
   
txtScId = PopUpValue2
txtschool.text = PopUpValue1  '& ", " & PopUpValue3
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""

fatchOrder

End If
End Sub

Private Sub txtschool_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   
   Screen.MousePointer = vbHourglass
  ''' tblNo = 9
  ''' frmSearchItem.Show
  
    searchType = "party"
    
    If txtNSCHNo.text = "" Then
    
        If cmbAgentName.text = "" Then
           popuplist_client "select ScName,ScId from billingSchoolQry group by ScName,ScId order by ScName", con
        Else
           popuplist_client "select ScName,ScId from billingSchoolQry where subledger='" & Me.customercode.text & "'  group by ScName,ScId order by ScName", con
        End If
    
    Else
    
        popuplist_client "SELECT SCName,ScId,AgentName FROM SchoolWiseBookReturnDet where EntryNo=" & txtNSCHNo.text & " group by ScName,ScId,AgentName order by ScName", con_LAST

    
    End If
    
    

    Screen.MousePointer = vbDefault
   
End If

End Sub
Private Sub weight_LostFocus()
weight = UCase(weight)
End Sub
Sub pp2()
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
    MaxLine = 50
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
        kkk.close
    End If
    CNSetup
    kkk.Open "select * from setup1", con, adOpenStatic, adLockReadOnly, adCmdText
    If flagyes = True Then

    
      If Not kkk.BOF Then
        Print #1, Chr(27) + Chr(15) + Chr(14)
        Print #1, Tab(T1); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname) '
        Print #1, Tab(T2 - 7); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
        Print #1, Tab(T3); Trim(kkk!phone1)
        Line = Line + 4
    End If
  
    Print #1, repli("-", 150)

  End If
    If rs1.State = 1 Then
        rs1.close
    End If
    rs1.Open "CREDITA", con, adOpenStatic, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.text) + "'"
        Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!subledger; Tab(T5); "Credit Note No. : "; Trim(rs1!invoiceNo); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!invoiceDate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.close
            End If
            kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE
                Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.: "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!BILTYDATE
                kkk.close
                Print #1,
                Print #1, Tab(70); "Demurrage.  : "; Trim(rs1!marka)
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
                kk.close
            End If
            kk.Open "select * from CREDITB where invoiceno=" + Trim(rs1!invoiceNo) + " and " & stringyear & " order by discount,printorder", con, adOpenStatic, adLockReadOnly, adCmdText
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
                        tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "' and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
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
                        tdata.close
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
                        tdata.Open "select sum(amount) from CREDITB where invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " and " & stringyear & " group by discount", con, adOpenStatic, adLockReadOnly, adCmdText
                        If Not tdata.BOF Then
                            Print #1, Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0), 2)), "0.00")), 12)
                            Print #1, Tab(T5); "Less Discount @ " + Trim(Format(Str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(Str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                            netamount = netamount + Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                        End If
                        tdata.close
                        Print #1, Tab(T7); repli("-", 22)
                Loop
            End If
           End If
           Print #1, Tab(T5 - 4); rsets(Trim(Str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
           Print #1, Tab(T6); repli("-", 22)
           If kk.State = 1 Then
                kk.close
           End If
           kk.Open "Select * from CREDITC where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenStatic, adLockReadOnly, adCmdText
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
           kk.close
           kk.Open "Select * from CREDITA where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenStatic, adLockReadOnly, adCmdText
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
            Print #1, Tab(0); repli("-", 120)
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            Dim LEFTM As Integer
            LEFTM = 5
            CNSetup
            tempdata.Open "setup1", con, adOpenStatic, adLockReadOnly, adCmdTable
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

Sub Backupprintinvoice()
Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = False
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = False
Me.Commandabandon.Enabled = True
Me.CommandPrint.Enabled = True
Me.Command1.Enabled = True
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
paperWidth = 150
MaxLine = 60
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
      kkk.close
End If
 CNSetup
 kkk.Open "select * from setup1", con, adOpenStatic, adLockReadOnly, adCmdText
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
     
     'line = line + 4
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


Print #1, Chr(27) + Chr(15) + Chr(14); Tab(25); dspace(Trim("CRDEIT NOTE ")); Chr(20); Tab(T4 + 6); IIf(Printheader = True, kkk!uptt, "")
If Printheader = True Then
   Print #1, Tab(T7 + 7); kkk!cst
Else
   Line = Line - 1
End If

Print #1, repli("-", 150)
Line = Line + 3
If rs1.State = 1 Then
   rs1.close
End If





If rs1.State = 1 Then
   rs1.close
End If
'Print #1, Chr(27) + Chr(14)
'line = line + 1
If rs1.State = 1 Then
    rs1.close
End If
rs1.Open "CreditA", con, adOpenStatic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.text) + "'"
If Not rs1.EOF Then
    Print #1, " To, S.L. Code : "; Tab(T1 + 8); Mid$(rs1!subledger, 1, 5); Tab(T5); "C/Note No. : "; Trim(rs1!invoiceNo); Tab(T8); "Dated     : "; rs1!invoiceDate
    kine = libe + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE; Tab(T5); "Bilty No.  : "; Trim(rs1!biltyno); Tab(T8); "Dated     : "; IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
        Print #1, Tab(3); kkk!address1; Tab(T5); "Freight    : "; Trim(rs1!freight); Tab(T8); "Bundle(s) :  "; Trim(rs1!bundles);
        Print #1, Tab(3); kkk!address2; Tab(T5); "Demurrage  : "; Trim(rs1!marka); Tab(T8); "Weight    :  "; Trim(rs1!weight)
        Print #1, Tab(3); kkk!address3; Chr(27) + Chr(72)
        kkk.close
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
        kk.close
    End If
    kk.Open "select * from CreditB where invoiceno=" + Trim(rs1!invoiceNo) + " and " & stringyear & " order by printorder", con, adOpenStatic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "' and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
                Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(Str(kk!QUANTITY)), 5); Tab(T6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!QUANTITY
                Line = Line + 1
                If Line > MaxLine Then
                    called1 = True
                    Line = 0
                    Print #1, Chr(12)
                    Line = Line + 1
                    GoTo header
printagain1:
                    called1 = False
               End If
                tdata.close
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
                  Line = Line + 1
                    GoTo header
printagain2:
                    called2 = False
                End If
                Print #1, Tab(T7); repli("-", 22)
                Line = Line + 1
                tdata.Open "select sum(amount) from CreditB where invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " and " & stringyear & " group by discount", con, adOpenStatic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                    Print #1, Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0), 2)), "0.00")), 12)
                    Print #1, Tab(T5); "Less Discount @ " + Trim(Format(Str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(Str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                    Print #1, Tab(T7); repli("-", 22)
                    Line = Line + 3
                    netamount = netamount + Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                End If
                tdata.close
                'Print #1, Tab(t7); repli("-", 22)
                'line = line + 1
                Loop
            End If
        End If
        Print #1, repli("-", 150)
        Print #1, Tab(T5 - 4); rsets(Trim(Str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
       'by vk Print #1, Tab(t8); repli("-", 22)
        Line = Line + 2
        If kk.State = 1 Then
             kk.close
        End If
        kk.Open "Select * from CreditC where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenStatic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5 + 10); Trim(kk!text) + "    " + Trim(Format(Str(Round(kk!rate, 2)), "0.00")); Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5 + 10); Trim(kk!text); Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T8); repli("-", 22)
            Print #1, Chr(27) + Chr(71); Tab(T5 - 10); "NET AMOUNT Cr. TO YOUR A/C :"; Tab(T8 + 6); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
            Line = Line + 2
            VNetamt = netamount
        End If
        kk.close
        kk.Open "Select * from CreditA where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenStatic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(T5 + 10); kk!txt1; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + Round(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(T5 + 10); kk!txt2; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + Round(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(T5 + 10); "BY BANK "; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - Round(kk!baa, 2)
             End If
        End If
        Print #1, Tab(T8); repli("-", 22)
        'Print #1, Tab(T5 + 10); Chr(27) + Chr(71); "BALANCE  : "; Tab(T8 + 6); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
       ' Print #1, Tab(T8); repli("-", 22)
       Line = Line + 1
       ' PRINT THE FOOTER IN INVOICE START
        Do While Line < 61
            Print #1, ""
            Line = Line + 1
        Loop
        Print #1, Tab(0); toword(Round(VNetamt, 2))
        Print #1, Tab(0); repli("-", 150)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        Dim LEFTM As Integer
        LEFTM = 5
        CNSetup
        tempdata.Open "setup1", con, adOpenStatic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(0); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!cname)) * 0.75))); "FOR " + Trim(tempdata!cname)
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
        'Me.Enabled = False
        
        PrintOption.Show
        
        'viewinvoice.Left = 0
        'viewinvoice.Top = 10
        'viewinvoice.Show
End Sub



