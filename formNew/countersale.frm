VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form countersale 
   ClientHeight    =   8412
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   12600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8412
   ScaleWidth      =   12600
   Begin VB.Frame panel 
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
      Height          =   8295
      Left            =   60
      TabIndex        =   22
      Top             =   0
      Width           =   12480
      Begin VB.TextBox txtchecked 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   6348
         MaxLength       =   100
         TabIndex        =   76
         Top             =   6624
         Width           =   540
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
         Left            =   144
         TabIndex        =   74
         Top             =   1296
         Visible         =   0   'False
         Width           =   2892
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
         Height          =   396
         Left            =   144
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   864
         Width           =   1308
      End
      Begin VB.TextBox txtScId 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   11325
         TabIndex        =   68
         Top             =   1365
         Width           =   660
      End
      Begin VB.TextBox txtAmtwords 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   765
         Left            =   9480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   7260
         Width           =   2895
      End
      Begin VB.ComboBox customercode 
         Height          =   1104
         Left            =   6480
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Top             =   288
         Visible         =   0   'False
         Width           =   5505
      End
      Begin VB.ComboBox Genledger 
         Height          =   315
         Left            =   11460
         Sorted          =   -1  'True
         TabIndex        =   43
         Top             =   1860
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   240
         ScaleHeight     =   792
         ScaleWidth      =   9060
         TabIndex        =   29
         Top             =   7125
         Width           =   9060
         Begin VB.CommandButton Commandhelp 
            Caption         =   "Help"
            Height          =   375
            Left            =   -825
            TabIndex        =   39
            Top             =   0
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&dit"
            Height          =   690
            Left            =   1050
            Picture         =   "countersale.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Height          =   690
            Left            =   2040
            Picture         =   "countersale.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   690
            Left            =   3045
            Picture         =   "countersale.frx":1026
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Enabled         =   0   'False
            Height          =   690
            Left            =   4035
            Picture         =   "countersale.frx":15B0
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   690
            Left            =   5040
            Picture         =   "countersale.frx":2194
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton CommandPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Enabled         =   0   'False
            Height          =   690
            Left            =   7035
            Picture         =   "countersale.frx":2D78
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   690
            Left            =   8040
            Picture         =   "countersale.frx":395C
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandadd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   690
            Left            =   45
            Picture         =   "countersale.frx":4540
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandprintnh 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N&HPrint"
            Enabled         =   0   'False
            Height          =   690
            Left            =   6045
            Picture         =   "countersale.frx":5124
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   45
            Width           =   975
         End
      End
      Begin VB.ComboBox Bookcode 
         Height          =   720
         Left            =   555
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   28
         Top             =   3105
         Width           =   2415
      End
      Begin VB.ComboBox Bookname 
         Height          =   912
         Left            =   3555
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   27
         Top             =   3105
         Width           =   2295
      End
      Begin VB.CommandButton Commandother 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&End Part"
         Height          =   600
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6210
         Width           =   930
      End
      Begin VB.CommandButton Commandall 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Books"
         Height          =   630
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   6195
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   696
         Left            =   156
         TabIndex        =   23
         Top             =   156
         Width           =   1320
         Begin VB.OptionButton Optioncash 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   135
            TabIndex        =   0
            Top             =   135
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton Optioncredit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Credit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   135
            TabIndex        =   1
            Top             =   405
            Width           =   840
         End
      End
      Begin VB.ComboBox Combosldistrictcode 
         Height          =   315
         Left            =   6495
         TabIndex        =   10
         Top             =   720
         Width           =   4845
      End
      Begin VB.ComboBox cmbdiscountcat 
         Height          =   315
         Left            =   11970
         TabIndex        =   7
         Top             =   990
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.ComboBox cmbAgentName 
         Height          =   315
         Left            =   990
         TabIndex        =   12
         Top             =   1365
         Width           =   3585
      End
      Begin VB.ComboBox cmbareaname 
         BackColor       =   &H80000003&
         Height          =   1104
         Left            =   6480
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Top             =   288
         Visible         =   0   'False
         Width           =   5490
      End
      Begin VB.ComboBox cmbtransportname 
         Height          =   315
         Left            =   1455
         TabIndex        =   16
         Top             =   1980
         Width           =   3120
      End
      Begin VB.ComboBox txtMark 
         Height          =   315
         ItemData        =   "countersale.frx":5D08
         Left            =   9000
         List            =   "countersale.frx":5D15
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1980
         Width           =   1170
      End
      Begin VB.ComboBox cboCatII 
         Height          =   315
         Left            =   11925
         TabIndex        =   8
         Top             =   945
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.ComboBox cboCatII1 
         Height          =   315
         Left            =   11880
         TabIndex        =   9
         Top             =   945
         Visible         =   0   'False
         Width           =   345
      End
      Begin MSMask.MaskEdBox textbox 
         Height          =   315
         Left            =   6495
         TabIndex        =   4
         Top             =   285
         Width           =   5490
         _ExtentX        =   9694
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3645
         Left            =   180
         TabIndex        =   26
         Top             =   2400
         Width           =   11595
         _ExtentX        =   20447
         _ExtentY        =   6435
         _Version        =   393216
         BackColorFixed  =   7917545
         ForeColorFixed  =   4210752
         GridColorFixed  =   12648447
         FillStyle       =   1
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox I_DTOB 
         Height          =   330
         Left            =   3420
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   593
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox bdated 
         Height          =   315
         Left            =   6225
         TabIndex        =   18
         Top             =   1980
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox bundles 
         Height          =   330
         Left            =   10200
         TabIndex        =   21
         Top             =   1980
         Width           =   1230
         _ExtentX        =   2180
         _ExtentY        =   593
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_OB 
         Height          =   285
         Left            =   11970
         TabIndex        =   11
         Top             =   990
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   593
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
         Height          =   330
         Left            =   3420
         TabIndex        =   3
         Top             =   615
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   593
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tempmeb 
         Height          =   285
         Left            =   915
         TabIndex        =   40
         Top             =   2595
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1672
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rate 
         Height          =   285
         Left            =   585
         TabIndex        =   41
         Top             =   3735
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3260
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox amount 
         Height          =   285
         Left            =   675
         TabIndex        =   42
         Top             =   3195
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3260
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_NO 
         Height          =   330
         Left            =   3420
         TabIndex        =   2
         Top             =   270
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   572
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox freight 
         Height          =   315
         Left            =   7635
         TabIndex        =   19
         Top             =   1980
         Width           =   1305
         _ExtentX        =   2307
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox biltno 
         Height          =   315
         Left            =   4665
         TabIndex        =   17
         Top             =   1980
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox station 
         Height          =   315
         Left            =   195
         TabIndex        =   15
         Top             =   1980
         Width           =   1245
         _ExtentX        =   2201
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtschool 
         Height          =   285
         Left            =   6495
         TabIndex        =   13
         Top             =   1365
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Checked :"
         Height          =   252
         Left            =   5580
         TabIndex        =   77
         Top             =   6660
         Width           =   684
      End
      Begin VB.Label lblState 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6510
         TabIndex        =   73
         Top             =   1080
         Width           =   4785
      End
      Begin VB.Label lblDId 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   11340
         TabIndex        =   72
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label8 
         Caption         =   "Amt in words :"
         Height          =   255
         Left            =   9480
         TabIndex        =   71
         Top             =   7020
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
         Left            =   4815
         TabIndex        =   70
         Top             =   1365
         Width           =   780
      End
      Begin VB.Label Label23 
         Caption         =   "Amt in words :"
         Height          =   255
         Left            =   1290
         TabIndex        =   69
         Top             =   7185
         Width           =   1035
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   915
         Left            =   195
         Top             =   7095
         Width           =   9165
      End
      Begin VB.Label tqu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4095
         TabIndex        =   66
         Top             =   6150
         Width           =   1185
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Quantity : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2970
         TabIndex        =   65
         Top             =   6150
         Width           =   1110
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Total Discount : "
         Height          =   255
         Left            =   7965
         TabIndex        =   64
         Top             =   3435
         Width           =   1290
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bundle(s):"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10110
         TabIndex        =   63
         Top             =   1725
         Width           =   1215
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2700
         TabIndex        =   62
         Top             =   945
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Gross Amount : "
         Height          =   255
         Left            =   6735
         TabIndex        =   61
         Top             =   3945
         Width           =   1200
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Net Amount : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8895
         TabIndex        =   60
         Top             =   6495
         Width           =   1200
      End
      Begin VB.Label label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Code : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4815
         TabIndex        =   59
         Top             =   315
         Width           =   1515
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Memo No. : "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2010
         TabIndex        =   58
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2700
         TabIndex        =   57
         Top             =   645
         Width           =   570
      End
      Begin VB.Label mga 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8880
         TabIndex        =   56
         Top             =   6195
         Width           =   1200
      End
      Begin VB.Label mna 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10185
         TabIndex        =   55
         Top             =   6495
         Width           =   1200
      End
      Begin VB.Label mgd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10185
         TabIndex        =   54
         Top             =   6195
         Width           =   1200
      End
      Begin VB.Label labelbybank 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4095
         TabIndex        =   53
         Top             =   6585
         Width           =   1200
      End
      Begin VB.Label labelbybanklbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "By Cash : "
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2985
         TabIndex        =   52
         Top             =   6585
         Width           =   1110
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "District Name"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4815
         TabIndex        =   51
         Top             =   810
         Width           =   1530
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Rep.Name :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   50
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transport"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1455
         TabIndex        =   49
         Top             =   1725
         Width           =   1935
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5955
         TabIndex        =   48
         Top             =   1725
         Width           =   1230
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Railway/Station : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   195
         TabIndex        =   47
         Top             =   1725
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
         Left            =   4605
         TabIndex        =   46
         Top             =   1725
         Width           =   1365
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Freight : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7560
         TabIndex        =   45
         Top             =   1725
         Width           =   1335
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mark"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9225
         TabIndex        =   44
         Top             =   1725
         Width           =   915
      End
   End
End
Attribute VB_Name = "countersale"
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
Dim Printheader As Boolean
Dim addoredit As Boolean
Dim emptyInv_bool As Boolean
Dim category As String
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
Me.Commandprintnh.Enabled = True
Dim called1, called2 As Boolean
Dim MaxLine As Integer
Dim netamount As Double
Dim FooterYes As Boolean
Dim totalquantity As Long
Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
Dim RS As ADODB.Recordset
Dim LEFTM As Integer
Set RS = New ADODB.Recordset

Dim mystr1 As String
mystr1 = ""
If Me.txtMark.text = "M" Then
mystr1 = "MOHKAMPUR"
ElseIf Me.txtMark.text = "W" Then
mystr1 = "W.K.ROAD"
ElseIf Me.txtMark.text = "U" Then
mystr1 = "UTSAV COMPLEX"
End If

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
paperWidth = 81
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
Pno = 1
LEFTM = 5
FooterYes = False
header:
    If kkk.State = 1 Then
          kkk.close
    End If
    CNSetup
    kkk.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If FooterYes = True Then
        If Line > MaxLine - 10 Then
            Do While Line < 61
                Print #1, ""
                Line = Line + 1
            Loop
        End If
        Line = 0
        LEFTM = 5
        Print #1, Tab(0); repli("-", 81)
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); kkk!COURT; Tab(50); "FOR " + Trim(kkk!cname)
        Print #1, ""
        Print #1, Tab(1); Chr(27) + Chr(71); "Continued on Page : " & Pno; Chr(27) + Chr(72)
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
     
     Print #1, Chr(27) + Chr(71); Chr(27) + Chr(77) + Chr(14)
     Print #1, Tab((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2); Chr(27) + Chr(77) + Chr(14); Trim(kkk!cname)
     Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(77); dspace(Trim(kkk!add1))
     Print #1, Tab((paperWidth - (Len(Trim(kkk!phone1)) * 2)) / 2); Trim(kkk!phone1) & "," & Trim(kkk!phone2)
     Line = Line + 7
   End If
Else
     
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, Chr(27) + Chr(77)
     Line = Line + 7
End If
If FooterYes = True Then
   Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72)
End If
Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "                     ", ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CASH MEMO")))) / 2 - 8); Chr(14); "***CASH MEMO***"; Chr(20)
Print #1, Tab(48); IIf(Printheader = True, kkk!uptt, "")
Line = Line + 2
If Printheader = True Then
   Line = Line + 1
   Print #1, Tab(48); kkk!cst
   
End If

Print #1, Tab(0); repli("-", 81)
Line = Line + 1

If rs1.State = 1 Then rs1.close
rs1.Open "select * from casha where " & stringyear & " and  invoiceno='" + Trim(Me.I_NO.text) + "'", con, adOpenStatic, adLockReadOnly
If Not rs1.EOF Then

Print #1, Chr(27) + Chr(71); "To, S.L. Code :"; Tab(19); IIf(Optioncash.value = True, "", Mid$(rs1!subledger, 1, 5)); Tab(38); "Cash Memo No.: "; Trim(rs1!invoiceNo); Tab(67); "Dt. : "; rs1!invoiceDate; Chr(27) + Chr(72);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & IIf(Optioncash.value = True, rs1!CASHPARTYNAME, kkk!DESCFORINVOICE)
        Print #1, Tab(5); IIf(IsNull(kkk!address1), "", kkk!address1); Tab(37); Chr(27) + Chr(71); "Bilty No     : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(68); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!address2), "", kkk!address2); Tab(37); Chr(27) + Chr(71); "Bundle(s)    : "; Chr(27) + Chr(72); Trim(rs1!bundles); Tab(64); Chr(27) + Chr(71); "Freight :"; Chr(27) + Chr(72); rs1!freight
        Print #1, Tab(5); IIf(IsNull(kkk!address3), "", kkk!address3); ; Tab(37); Chr(27) + Chr(71); "Agent Name   : "; Chr(27) + Chr(72); Trim(rs1!agentname)
        Print #1, Tab(5); "Station : " + IIf(IsNull(rs1!station), "", rs1!station) + " " + IIf(IsNull(rs1!transportname), "", rs1!transportname); Tab(73); Chr(27) + Chr(71); "(" & txtMark & ")"; Chr(27) + Chr(72)
        kkk.close
        
        Print #1, Chr(27) + Chr(71); repli("-", 81)
        Print #1, Tab(0); "S.No."; Tab(10); "Book Description"; Tab(44); "Qty."; Tab(52); "Rate"; Tab(61); "Amount"; Tab(71); "Net Amount"
        Print #1, repli("-", 81); Chr(27) + Chr(72)
        Line = Line + 8
    End If
    If called1 Then
        GoTo printagain1
    End If
    If called2 Then
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.close
    kk.Open "select * from CASHB where " & stringyear & " and  invoiceno=" + Trim(rs1!invoiceNo) + " order by printorder,sno ", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
                Print #1, Tab(0); rsets(Trim(Str(sno)), 4); Tab(6); Trim(tdata!Bookname); Tab(41); rsets(Trim(Str(kk!QUANTITY)), 5); Tab(48); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(56); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
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
            If Line > MaxLine - 10 Then
                    called2 = True
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
                    
                    
printagain2:
                    called2 = False
                End If
                Print #1, Tab(57); repli("-", 12)
                Line = Line + 1
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CashB where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and printorder =" + Trim(Str(cdiscount)) + " group by printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(56); rsets(Trim(Format(Str(tdata(0)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(vdis), "0.00")) + " %"; Tab(56); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(69); rsets(Trim(Format(Str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(57); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.close
             Loop
         End If
    End If
    Print #1, repli("-", 81)
    Print #1, Tab(39); rsets(Trim(Str(totalquantity)), 7); Tab(69); rsets(Trim(Format(Str(netamount), "0.00")), 12)
    Line = Line + 2
    If kk.State = 1 Then kk.close
    kk.Open "Select * from CASHC where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
       Do While Not kk.EOF
             If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(48); Trim(kk!text) + "    " + Trim(Format(Str(kk!rate), "0.00")); Tab(69); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(48); Trim(kk!text); Tab(69); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
          
        End If
        Print #1, Tab(69); repli("-", 12)
        Print #1, Chr(27) + Chr(71); Tab(49); "NET AMOUNT  : "; Tab(70); rsets(Trim(Format(Str(netamount), "0.00")), 12); Chr(27) + Chr(72)
        VNetamt = netamount
        Line = Line + 2
        kk.close
        kk.Open "Select * from CASHA where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(48); kk!txt1 & "    :"; Tab(69); rsets(Trim(Format(Str(Abs(kk!txt1a)), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + kk!txt1a
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(48); kk!txt2 & " :"; Tab(69); rsets(Trim(Format(Str(Abs(kk!txt2a)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + kk!txt2a
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(48); "CASH RECD.  :"; Tab(69); rsets(Trim(Format(Str(Abs(kk!baa)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - kk!baa
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(69); repli("-", 12)
                 Print #1, Tab(48); Chr(27) + Chr(71); "BALANCE     : "; Tab(70); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
                 Line = Line + 2
              End If
        End If
        Print #1, Tab(69); repli("-", 12);
          Line = Line + 1
        'PRINT THE FOOTER IN INVOICE START
        Do While Line < 61
            Print #1, ""
            Line = Line + 1
        Loop
        Print #1, Tab(0); Chr(27) + Chr(71); toword(Round(VNetamt, 2)); Chr(27) + Chr(72)
        Print #1, Tab(0); repli("-", 81)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        CNSetup
        tempdata.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(50); "FOR " + Trim(tempdata!cname)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
       ' Print #1, ""
        'PRINT THE FOOTER IN INVOICE END
        Close #1
        PrintOption.Show
        
End Sub

Sub invoicecalc()
'OTHERCASH.calc
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     mna.Caption = Format(Round((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
End Sub
Sub invoiceabandon()
        
        txtchecked.text = ""

        Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
        Me.Commandprintnh.Enabled = True
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        If kk.State = 1 Then
             kk.close
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
        For I = 1 To maxrow
           grid1.Row = I
            For J = 1 To 8
                grid1.Col = J
               grid1.text = ""
           Next
        Next
        I_DTOB = "__/__/____"
        bdated = "__/__/____"
        tqu.Caption = ""
        mga.Caption = ""
        mgd.Caption = ""
        mna.Caption = ""
        lblState.Caption = ""
        labelbybank.Caption = ""
        cmbAgentName.text = "."
        lblDId.Caption = ""
        maxrow = 0
        addoredit = False
        Unload OTHERCASH
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
    RRR = grid1.Row
    CCC = grid1.Col
    grid1.Row = lastrow
    grid1.Col = lastcol
    mprevcol = grid1.Col
    Select Case grid1.Col
            Case 1
                grid1.text = tempmeb.text
                '/*************************
                'If RS.State = 1 Then
                '    RS.close
                'End If
                'RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                
                Set RS = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(grid1.text) & "'")
                Row = grid1.Row
                Col = grid1.Col
                If Trim(grid1.text) <> "" Then
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
                            grid1.text = RS(0)
                            grid1.Col = 2
                            grid1.text = RS(1)
                         '   If Not edit Then
                                grid1.Col = 3
                                If Trim(grid1.text) = "" Then
                                    grid1.text = 0
                                End If
                                q = Val(grid1.text)
                                grid1.Col = 5
                                If Trim(grid1.text) = "" Then
                                grid1.text = Format(RS(3), "0.00")            'rs(3)
                                r = RS(3)
                                End If
                                '/******************
                                
'------------------------------------------
                            category = returnCategory(Trim(RS(2)))
                            If Optioncash.value = True Then
                            
                            If category = "C1" Then
                               Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                            ElseIf category = "C2" Then
                               Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                            ElseIf category = "C3" Then
                               Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII1.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                            End If
                            
                            Else
                            
                            If category = "C1" Then
                               Set kk = con.Execute("select DISCATEGORY from sledger where " & stringyear & " and subledger='" + Trim(customercode.text) + "'")
                            ElseIf category = "C2" Then
                               Set kk = con.Execute("select CATEGORY2 from sledger where " & stringyear & " and subledger='" + Trim(customercode.text) + "'")
                            ElseIf category = "C3" Then
                               Set kk = con.Execute("select CATEGORY3 from sledger where " & stringyear & " and subledger='" + Trim(customercode.text) + "'")
                            End If
                                
                                
                            End If
'-----------------------------------
                                
                                
                                grid1.Col = 6
                                
                                 If kk.BOF Then
                                             GoTo abc
                                 End If
                                       
                                
                                If grid1.text = "" And addmode = True Then
                                    If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        
                                kk.close
                                If category = "C1" Then
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                ElseIf category = "C2" Then
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                ElseIf category = "C3" Then
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII1.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                End If
                                        
                                        
                                        
                                        If kk.BOF Then
                                             GoTo abc
                                        End If
                                        
                                        grid1.Col = 4
                                        grid1.text = Format(kk(0), "0.00")
                                        grid1.Col = 6
                                        grid1.text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = RS(3)
                                     Else
abc:
                                        grid1.Col = 4
                                        grid1.text = Format(RS(4), "0.00")
                                        grid1.Col = 6
                                        grid1.text = Format(RS(4), "0.00")
                                        D = RS(4)
                                    End If
                                    
                                    
                                    'New Code For Sub Group
                                    If IsEmpty(tempstr) Then tempstr = ""
                                    s_ = tempstr
                                    
                                    If rs1.State = 1 Then rs1.close
                                    rs1.Open "select GROUPCODE_sub from books where bookcode='" & RS(0) & "'", con
                                    If rs1.EOF = False Then
                                    If (Not IsNull(rs1!GROUPCODE_sub) And Len(rs1!GROUPCODE_sub) > 0) Then
                                        D = ReturnDiscount("" & category, "" & s_, Trim(rs1(0)))
                                        If D > 0 Then
                                            grid1.Col = 4
                                            grid1.text = Format(D, "    0.00")
                                            grid1.Col = 6
                                            grid1.text = Format(D, "0.00")
                                            r = RS(3)
                                        End If
                                    End If
                                    End If
                                    'End Code For Sub Group
                                    
                                   '===================Serise Wise============================
                                     D = ReturnDiscountNew(RS(0), Trim(customercode.text), txtScId.text)
                                     If D > 0 Then
                                        grid1.Col = 4
                                        grid1.text = Format(D, "    0.00")
                                        grid1.Col = 6
                                        grid1.text = Format(D, "0.00")
                                        r = RS(3)
                                    End If
                                  '==========================================================
                                    
                                    
                                    
                                    
                                    
                                    grid1.Col = 7
                                    grid1.text = Format(Round(q * r, 2), "0.00")
                                    grid1.Col = 8
                                    grid1.text = Format(Round((q * r) * (D / 100), 2), "0.00")
                            Else
                                                    
                            
                            If grid1.text = "" And addmode = False Then
                                If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    
                                    
                                kk.close
                                If category = "C1" Then
                                    
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                
                                ElseIf category = "C2" Then
                                    
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                
                                ElseIf category = "C3" Then
                                    
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII1.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                        
                                 End If
                                    
                                    
                                    
                                    
                                    
                                    grid1.Col = 4
                                    If kk.BOF Then
                                        GoTo abc
                                    End If
                                    grid1.text = Format(kk(0), "0.00")
                                    grid1.Col = 6
                                    grid1.text = Format(kk(0), "0.00")
                                    D = kk(0)
                                    r = RS(3)
                            
                            
                                 
                                 
                                 
                                 
                                    'New Code For Sub Group
                                    If IsEmpty(tempstr) Then tempstr = ""
                                    s_ = tempstr
                                    
                                    If rs1.State = 1 Then rs1.close
                                    rs1.Open "select GROUPCODE_sub from books where bookcode='" & RS(0) & "'", con
                                    If rs1.EOF = False Then
                                    If (Not IsNull(rs1!GROUPCODE_sub) And Len(rs1!GROUPCODE_sub) > 0) Then
                                        D = ReturnDiscount("" & category, "" & s_, Trim(rs1(0)))
                                        If D > 0 Then
                                            grid1.Col = 4
                                            grid1.text = Format(D, "    0.00")
                                            grid1.Col = 6
                                            grid1.text = Format(D, "0.00")
                                             r = RS(3)
                                        End If
                                    End If
                                    End If
                                    'End Code For Sub Group
                                    
                                  '===================Serise Wise============================
                                     D = ReturnDiscountNew(RS(0), Trim(customercode.text), txtScId.text)
                                     If D > 0 Then
                                        grid1.Col = 4
                                        grid1.text = Format(D, "    0.00")
                                        grid1.Col = 6
                                        grid1.text = Format(D, "0.00")
                                        r = RS(3)
                                    End If
                                  '==========================================================
                                 
                                 
                            
                            
                            
                                End If
                            End If
                            End If
                            grid1.Col = Col
                            RS.close
                        End If
                    End If
                End If
            Case 3, 5, 6
                If grid1.Col <> 3 Then
                    grid1.text = Format(Trim(tempmeb.text), "0.00")
                Else
                    grid1.text = Format(Trim(tempmeb.text), "0")
                End If
                If Trim(grid1.text) = "" Then
                    grid1.text = 0
                End If
                Row = grid1.Row
                Col = grid1.Col
                grid1.Col = 3
                q = Val(Trim(grid1.text))
                grid1.Col = 5
                r = Val(Trim(grid1.text))
                grid1.Col = 6
                D = Val(Trim(grid1.text))
                grid1.Col = 7
                grid1.text = Format(Round(q * r, 2), "0.00")
                grid1.Col = 8
                grid1.text = Format(Round((q * r) * (D / 100), 2), "0.00")
                grid1.Col = Col
            Case 4
                grid1.text = tempmeb.text
                If Trim(grid1.text) = "" Then
                    grid1.text = 0
                End If
        End Select
        Row = grid1.Row
        Col = grid1.Col
        
        
        
        If (Col = 1 Or Col = 6 Or Col = 3) Then
        
            totalamount = 0
            totaldiscount = 0
            For I = 1 To maxrow
                grid1.Row = I
                grid1.Col = 7
                totalamount = totalamount + Val(Trim(grid1.text))
                grid1.Col = 8
                totaldiscount = totaldiscount + Val(Trim(grid1.text))
            Next
            invoicecalc
            Me.tqu.Caption = ""
            For I = 1 To maxrow
                grid1.Col = 3
                grid1.Row = I
                Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(grid1.text))
            Next
        
        End If
        
        
        grid1.Row = RRR
        grid1.Col = CCC
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
        mprevcol = grid1.Col
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        Select Case grid1.Col
            Case 2
                Dim Row, Col As Integer
                Row = grid1.Row
                Col = grid1.Col
                If Trim(Bookname.text) = "" Then
                    grid1.Col = 1
                    If Trim(grid1.text) = "" Then
                        grid1.text = Bookname.text
                           Bookname.SetFocus
  '********* vk
                          
                          
                          If Trim(grid1.text) = "" And Row = 1 Then
                                 grid1.Col = 2
                                 grid1.text = ""
                                 If Trim(grid1.text) = "" Then
                                           
                                          grid1.Col = 1
                                          Bookname.SetFocus
                                          grid1.SetFocus
                                       Exit Sub
                                 End If
                           End If
              '********
                           Commandother.SetFocus
                           'station.SetFocus
                           
                        Exit Sub
                    End If
                End If
                grid1.Row = Row
                grid1.Col = Col
                grid1.text = Bookname.text
                '/*************************
                If RS.State = 1 Then
                    RS.close
                End If
                RS.Open "select * from books where  " & stringyear & " and  bookname='" + Trim(grid1.text) + "'", con, adOpenDynamic, adLockReadOnly
                Row = grid1.Row
                Col = grid1.Col
                If Trim(grid1.text) <> "" Then
                    If Not RS.BOF Then
                        'RS.MoveFirst
                        'RS.Find "bookname='" + Trim(Grid1.Text) + "'"
                        If RS.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            RS.close
                            Exit Sub
                        Else
                            
                            grid1.Col = 1
                            grid1.text = RS(0)
                            grid1.Col = 2
                            grid1.text = RS(1)
                        '    If Not edit Then
                                 grid1.Col = 3
                                If Trim(grid1.text) = "" Then
                                        grid1.text = 0
                                End If
                                q = Val(grid1.text)
                                grid1.Col = 5
                                grid1.text = Format(RS(3), "0.00")
                                r = RS(3)
                                '/******************
                                
 
'------------------------------------------
                            category = returnCategory(Trim(RS(2)))
                            If Optioncash.value = True Then
                            
                                If category = "C1" Then
                                   Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                ElseIf category = "C2" Then
                                   Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                ElseIf category = "C3" Then
                                   Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII1.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                End If
                                
                            
                            Else
                            
                            If category = "C1" Then
                               Set kk = con.Execute("select DISCATEGORY from sledger where " & stringyear & " and subledger='" + Trim(customercode.text) + "'")
                            ElseIf category = "C2" Then
                               Set kk = con.Execute("select CATEGORY2 from sledger where " & stringyear & " and subledger='" + Trim(customercode.text) + "'")
                            ElseIf category = "C3" Then
                               Set kk = con.Execute("select CATEGORY3 from sledger where " & stringyear & " and subledger='" + Trim(customercode.text) + "'")
                            End If
                                
                                
                            End If
'-----------------------------------
 
                                   If kk.BOF Then
                                      GoTo abc
                                   End If
 
 
                                
                                grid1.Col = 6
                                If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    kk.close
                                If category = "C1" Then
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                
                                ElseIf category = "C2" Then
                                
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                   
                                    End If
                                        
                                ElseIf category = "C3" Then
                                
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII1.text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                End If
                                
                                   
                                If kk.BOF Then
                                      GoTo abc
                                   End If
                                    grid1.Col = 4
                                    grid1.text = Format(kk(0), "0.00")
                                    grid1.Col = 6
                                    grid1.text = Format(kk(0), "0.00")
                                    D = kk(0)
                                Else
abc:
                                    grid1.Col = 4
                                    grid1.text = Format(RS(4), "0.00")
                                    grid1.Col = 6
                                    grid1.text = Format(RS(4), "0.00")
                                    D = RS(4)
                                End If
                                
                                
                                
                                
                                    'New Code For Sub Group
                                    If IsEmpty(tempstr) Then tempstr = ""
                                    s_ = tempstr
                                    
                                    If rs1.State = 1 Then rs1.close
                                    rs1.Open "select GROUPCODE_sub from books where bookcode='" & RS(0) & "'", con
                                    If rs1.EOF = False Then
                                    If (Not IsNull(rs1!GROUPCODE_sub) And Len(rs1!GROUPCODE_sub) > 0) Then
                                        D = ReturnDiscount("" & category, "" & s_, Trim(rs1(0)))
                                        If D > 0 Then
                                            grid1.Col = 4
                                            grid1.text = Format(D, "    0.00")
                                            grid1.Col = 6
                                            grid1.text = Format(D, "0.00")
                                            r = RS(3)
                                        End If
                                    End If
                                    End If
                                    'End Code For Sub Group
                                
                                
                                
                                
                                grid1.Col = 7
                                grid1.text = Round(q * r, 2)
                                grid1.Col = 8
                                grid1.text = Round((q * r) * (D / 100), 2)
                         '   End If
                            grid1.Col = Col
                            RS.close
                        End If
                    End If
                End If
        End Select
        Row = grid1.Row
        Col = grid1.Col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            grid1.Row = I
            grid1.Col = 7
            totalamount = totalamount + Val(Trim(grid1.text))
            grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.text))
        Next
        invoicecalc
        grid1.Row = Row
        grid1.Col = Col
        Select Case grid1.Col
            Case 1
                grid1.Col = 3
                grid1.SetFocus
                Grid1_Click
            Case 2
                grid1.Col = 3
                grid1.SetFocus
                Grid1_Click
            Case 3, 4, 5
                grid1.Col = grid1.Col + 1
                grid1.SetFocus
                Grid1_Click
            Case 6
                grid1.Col = 1
                grid1.Row = grid1.Row + 1
                grid1.SetFocus
                Grid1_Click
        End Select
    End If
End Sub
Private Sub Bookname_LostFocus()
    Bookname.Visible = False
End Sub

Private Sub bundles_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If Trim(Me.customercode.text) <> "" Then
            Me.grid1.Col = 1
            Me.grid1.Row = 1
            Me.grid1.SetFocus
            Me.Grid1_Click
        Else
            I_NO.SetFocus
            'Me.textbox.SetFocus
            
            'Me.customercode.SetFocus
        End If
    End If
End Sub

Private Sub bundles_LostFocus()
bundles = UCase(bundles)
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

Private Sub cmbareaname_LostFocus()
  Me.textbox.text = Me.textbox.text + ", " + cmbareaname.text
  cmbareaname.Visible = False
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
Set RS = con.Execute("exec searchList 'cash'")

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

Private Sub Combosldistrictcode_Click()
'lblDId = Combosldistrictcode.ItemData(Combosldistrictcode.ListIndex)
findDisId
End Sub
Sub findDisId()
   If rs1.State = 1 Then rs1.close
   rs1.Open "Select * from districtView where District = '" & Combosldistrictcode.text & "'", CON_blue, adOpenStatic, adLockReadOnly
   If rs1.RecordCount > 0 Then
      lblDId.Caption = rs1!DistrictID
      lblState.Caption = rs1!State
   End If

End Sub
Private Sub Combosldistrictcode_LostFocus()

If Combosldistrictcode.text = "" Then
   Combosldistrictcode.SetFocus
   Exit Sub
End If
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

If Combosldistrictcode.text <> "" Then
   rs1.Open "Select * from districtView where District = '" & Combosldistrictcode.text & "'", CON_blue, adOpenStatic, adLockReadOnly
   If rs1.RecordCount <= 0 Then
      MsgBox "Please Select valid district.."
      Combosldistrictcode.SetFocus
   End If
End If

findDisId

''Set rs1 = New ADODB.Recordset
''If Combosldistrictcode.Text <> "" And addmode = True Then
''   rs1.Open "Select * from districtView where District = '" & Combosldistrictcode.Text & "'", con, adOpenStatic, adLockReadOnly
''   If rs1.RecordCount > 0 Then
''      Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
''   End If
''End If





End Sub

Private Sub Command1_Click()
Printheader = True
   printinvoice
End Sub

Private Sub Commandprintnh_Click()

    printch = "CASHA"
    ino = I_NO
    printch1 = "INVOICENO"

s1 = "200"

Printheader = False
printinvoice
End Sub

Private Sub Commandabandon_Click()
invoiceabandon
Me.Commandall.Enabled = False
Me.Commandother.Enabled = False
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub
Private Sub Commandadd_Click()
On Error Resume Next
    invoiceabandon
    Dim RS As ADODB.Recordset
    addoredit = True
    addmode = True
    Edit = False
    Set RS = New ADODB.Recordset
    Dim TEMPNUM As Integer
    
    If Edit = False Then
    'If CON.Execute("Select max(invoiceno) from CASHA")(0) >= Val(Trim(Me.I_NO.Text)) Then
         RS.Open "Select max(invoiceno) from CASHA where " & stringyear, con, adOpenDynamic, adLockOptimistic
         If IsNull(RS(0)) Then
           Me.I_NO.text = 1
         Else
           Me.I_NO.text = RS(0) + 1
         End If
         
         
         RS.update
         RS.close
        
     'End If
    End If
    
    Dim ctl As Control
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton Then
           ctl.Enabled = True
        End If
    Next
    
    Picture5.Enabled = True
    Commandother.Enabled = True
    Commandadd.Enabled = False
    Commanddelete.Enabled = False
    Commandedit.Enabled = False
    CommandPrint.Enabled = False
    Commandprintnh.Enabled = False
    Commandall.Enabled = True
    Commandsave.Enabled = False
    Commandsearch.Enabled = False
    Commandsave.Enabled = flase
    
    grid1.Enabled = True
    Me.customercode.Enabled = True
    Optioncash.SetFocus
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
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
    
    
    
    grid1.rows = 1
    grid1.rows = 2
    grid1.Col = 1
    grid1.Row = 1
    If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select * from books where " & stringyear & " order by BOOKCODE", con, adOpenDynamic, adLockReadOnly, adCmdText
    Row = grid1.Row
    Col = grid1.Col
    If Not RS.BOF Then
        RS.MoveFirst
        Do While Not RS.EOF
            grid1.Col = 1
            grid1.text = RS(0)
            grid1.Col = 2
            grid1.text = RS(1)
            grid1.Col = 3
            If Trim(grid1.text) = "" Then
                grid1.text = Val(myvalue)
            End If
            q = Val(grid1.text)
            grid1.Col = 5
            grid1.text = Format(RS(3), "0.00")            'rs(3)
            r = RS(3)
            '/******************
            
            
           category = returnCategory(Trim(RS(2)))
           If category = "C1" Then
            Set kk = con.Execute("select DISCATEGORY from sledger where " & stringyear & " and subledger='" + Trim(customercode.text) + "'")
           ElseIf category = "C2" Then
            Set kk = con.Execute("select Category2 from sledger where " & stringyear & " and subledger='" + Trim(customercode.text) + "'")
           ElseIf category = "C3" Then
            Set kk = con.Execute("select Category3 from sledger where " & stringyear & " and subledger='" + Trim(customercode.text) + "'")
           End If

            
            grid1.Col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.close
                Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
                grid1.Col = 4
                If kk.BOF Then
                    GoTo abc
                End If
                grid1.text = Format(kk(0), "0.00")
                grid1.Col = 6
                grid1.text = Format(kk(0), "0.00")
                D = kk(0)
            Else
abc:
                grid1.Col = 4
                grid1.text = Format(RS(4), "0.00")
                grid1.Col = 6
                grid1.text = Format(RS(4), "0.00")
                D = RS(4)
            End If
            grid1.Col = 7
            grid1.text = Format(Round(q * r, 2), "0.00")
            grid1.Col = 8
            grid1.text = Format(Round((q * r) * (D / 100), 2), "0.00")
            If Not RS.EOF Then
                grid1.rows = grid1.rows + 1
                grid1.Row = grid1.Row + 1
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
    For I = 1 To grid1.rows - 1
            grid1.Row = I
            grid1.Col = 7
            totalamount = totalamount + Val(Trim(grid1.text))
            grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.text))
            grid1.Col = 3
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(grid1.text))
     Next
     maxrow = grid1.rows - 1
Else
'Grid1_Click
Exit Sub
End If

invoicecalc
txtMark.ListIndex = 0

End Sub

Private Sub Commanddelete_Click()

On Error GoTo Del


Dim rs_h As New ADODB.Recordset
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset


If rs1.State = 1 Then rs1.close
rs1.Open "select * from casha where " & stringyear & " and invoiceno=" & I_NO.text & "", con
If rs1.EOF = False Then
   
If rs_h.State = 1 Then rs_h.close
rs_h.Open "select * from casha where " & stringyear & " and invoiceno=" & I_NO.text & "", con
   If rs1!bAuthorized = True Then
       MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
       Exit Sub
   End If
'End If
   
End If
  



If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
        
        
Else
                
                
                If (AuditTrail = "y") Then
                
                If (txtchecked.text = "y") Then
                
                    actionType_ = "Delete"
                    vtype1_ = "CM"
                    vtypeNew = "CM"
                    vdate_ = Trim(i_dt.text)
                    vno_ = Trim(I_NO.text)
                    
                    frmAuditTrailLog_Rem.Show 1
                    
                 End If
                
                End If



                
                
                con.Execute ("delete  from CASHA where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
                con.Execute ("delete  from CASHB where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
                con.Execute ("delete  from CASHC where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
                con.Execute ("delete  from CASHB_Free where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
                invoiceabandon
End If



Exit Sub
Del:
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub Commandedit_Click()
   If I_NO.text <> "" Then
  '  Exit Sub
    End If
    Commandadd.Enabled = False
    Me.Commandedit.Enabled = False
    Picture5.Enabled = True
    'Commandother.Enabled = True
    Commandadd.Enabled = False
    Commandedit.Enabled = False
    Commandall.Enabled = True
    Commandsave.Enabled = False
    Commanddelete.Enabled = False
    Commandsearch.Enabled = False
    CommandPrint.Enabled = False
     Commandprintnh.Enabled = False
    grid1.Enabled = True
    Commandall.Enabled = False
    Me.customercode.Enabled = True
    Edit = True
    I_NO_LostFocus
    i_dt.Enabled = True
    i_dt.SetFocus
    
    ' CASHCTMP creation start
    DoEvents
    con.Execute "Delete  from CASHCTMP where  " & stringyear & " and INVOICENO = " + Trim(I_NO.text) + ""
    DoEvents
    con.Execute ("insert into CASHCTMP(INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate,AMOUNT,DEBITORCREDIT,[TEXT],RYN,Fyear,setupid)  select INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate,AMOUNT,DEBITORCREDIT,[TEXT],RYN,Fyear,setupid from CASHC where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
    'invoicetmp creation end
    
    Dim KS As Long
    KS = 1
    For L = 1 To 15000
      PP = 0
    Next L
    DoEvents
    On Error Resume Next
    addoredit = False
    HIT
    VB.Screen.ActiveForm.ActiveControl.SelLength = 2
    OTHERCASH.top = 0
    OTHERCASH.Left = 0
    OTHERCASH.Visible = False
    
    
    Dim ctl As Control
    For Each ctl In Me.Controls
    If Not TypeOf ctl Is CommandButton Then
        ctl.Enabled = True
    End If
    Next

SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
End Sub
Private Sub Commandother_Click()

Commandsave.Enabled = True
searchForm = "cash"
frmEndPartTrans.Show
frmEndPartTrans.Refresh

    
End Sub
Private Sub CommandPrint_Click()
   
   
printch = "casha"
ino = I_NO
s1 = "200"
printch1 = "INVOICENO"
Printheader = True
printinvoice

End Sub
Private Sub CommandReturn_Click()

Unload Me
addoredit = False

''MainMenu.Toolbar1.Visible = True
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

 
 If rs1.State = 1 Then rs1.close
 rs1.Open "select * from casha where invoiceno=" & I_NO.text & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
 If rs1.EOF = False Then
    
    If rs_h.State = 1 Then rs_h.close
    rs_h.Open "select * from casha where invoiceno=" & I_NO.text & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
       If rs1!bAuthorized = True Then
           MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
           Exit Sub
       End If
    
 End If

  
  
  Dim RS As ADODB.Recordset
  Set RS = New ADODB.Recordset

  
  
  If Edit = False And addmode = False Then
      Me.Commandsave.Enabled = False
      Exit Sub
    End If


If Optioncash = True Then
    
'    If Trim(Combosldistrictcode) = "" Then
'        MsgBox "Please Enter District"
'        Exit Sub
'    End If
    If Val(Trim(Me.mna.Caption)) <> Val(Trim(frmEndPartTrans.T3TEXT.text)) Then
      MsgBox "In This Bill  Netamount and Cash Reciept Are Not Equal." + Chr(13) + "Please Select Ctedit Option For Part Cash Memo"
      Exit Sub
    End If
   

    
    
Else
   If Trim(cmbAgentName.text) = "" Then
        MsgBox "Please Enter Agent Name.."
        cmbAgentName.SetFocus
        Exit Sub
   End If
End If











Dim SAVED As Boolean
Dim LAMOUNT As Double
If MsgBox("Do you want to save it now ?", vbYesNo) = vbNo Then
    Exit Sub
End If
SAVED = False
grid1.Row = 1
grid1.Col = 1
If Trim(grid1.text) = "" Then
   MsgBox "Please Enter item.... "
   Exit Sub
End If


'----------------------------------------------------------------
'I_NO.Text = 852
If Edit = False Then
   If check_Duplikate("casha", I_NO.text) = True Then
      MsgBox "This  Inv. Number Already Exist ..", vbCritical
      Exit Sub
   End If
End If
'----------------------------------------------------------------


If Trim(I_NO.text) <> "" And Trim(i_dt.text) <> "" And Trim(customercode.text) <> "" Then
   
   If Edit Then
      con.Execute ("delete  from CASHA where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
      con.Execute ("delete  from CASHB where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
      con.Execute ("delete  from CASHC where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
      con.Execute ("delete  from CashRegister where cmno = " + Trim(I_NO.text))
      con.Execute ("delete  from CASHB_Free where " & stringyear & " and INVOICENO = " + Trim(I_NO.text))
   End If
   
   If RS.State = 1 Then RS.close
   LAMOUNT = 0
   
 If Edit = False Then
 If (I_OB <> "" And txtMark <> "") Then
 '    Party_Remove_FromOrder Trim(Me.customercode.Text), txtMark, Trim(I_OB)
 End If
 End If
 
   RS.Open "select * from CASHA where  " & stringyear & " and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
   If Not Edit Then
again:
      If con.Execute("Select max(invoiceno) from CASHA")(0) >= Val(Trim(Me.I_NO.text)) Then
      
      End If
   End If
   RS.AddNew
   
   If (AuditTrail = "y") Then
   RS!Checked_YesNo = Checked_YesNo
   End If
   
   RS!invoiceNo = Val(Me.I_NO.text)
   RS!invoiceDate = Me.i_dt.text
   RS!Genledger = Trim(Me.Genledger.text)
   
   RS!subledger = Trim(Me.customercode.text)
   
   RS!orderby = Trim(Me.I_OB.text)
   If Trim(Me.I_DTOB) = Trim("__/__/____") Then
     
      RS!ORDERDATE = Null
   Else
      RS!ORDERDATE = Trim(Me.I_DTOB.text)
   End If
   'rs!marka = Trim(Me.marka.Text)
   RS!bundles = Trim(Me.bundles)
   RS!Godown = txtMark.text
   'rs!through1 = Trim(Me.through1.Text)
   'If Trim(Me.through1.Text) = "" Then
   '   rs!through1 = " "
   'End If
   RS!station = Trim(Me.station.text)
   RS!biltyno = Trim(Me.biltno.text)
   If Trim(Me.bdated) = Trim("__/__/____") Then
      RS!BILTYDATE = Null
      'rs!BILTYDATE = Date
   Else
     RS!BILTYDATE = Me.bdated & ""
   End If
   RS!transportname = Trim(Me.cmbtransportname.text)
   RS!freight = Me.freight & ""
  ' rs!weight = Me.weight & ""
   RS!netamount = Round(Val(Trim(Me.mna.Caption)), 2)
   RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
   RS!txt1 = Trim(frmEndPartTrans.T1TEXT.text)
   RS!txt1a = Val(Trim(frmEndPartTrans.T1.text))
   RS!txt2 = Trim(frmEndPartTrans.T2TEXT.text)
   RS!txt2a = Val(Trim(frmEndPartTrans.T2.text))
   'rs!baa = Val(Trim(frmEndPartTrans.T3TEXT.Text))
   
   RS!baa = Val(Trim(frmEndPartTrans.T3TEXT.text))
   RS!baa = Val(Trim(labelbybank.Caption))
   
   RS!District = Combosldistrictcode.text
   RS!CASHPARTYNAME = textbox.text
   RS!agentname = cmbAgentName.text
   RS!discat = cmbdiscountcat.text
   RS!discatII = cboCatII.text
   RS!discatIII = cboCatII1.text
   RS!cityId = lblDId.Caption
   RS!states = lblState.Caption
   
err1:
   If Not Edit Then
      If con.Execute("Select max(invoiceno) from CASHA")(0) >= Val(Trim(Me.I_NO.text)) Then
         'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
         'rs!INVOICENO = Val(Me.I_NO.Text)
         On Error GoTo err1
      End If
   End If
               
   RS!Amtwords = txtAmtwords.text
   RS!scname = Trim(txtSchool)
   RS!scid = Trim(txtScId.text)
   RS!fyear = session
   RS!setupid = setupid

   RS.update
   On Error GoTo 0
   RS.close
   RS.Open "select * from CASHB where " & stringyear & "  and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
   Dim I As Integer
   RRRR = grid1.Row
   CCCC = grid1.Col
   For I = 1 To maxrow
       grid1.Row = I
       grid1.Col = 1
       If Trim(grid1.text) <> "" Then
          grid1.Col = 3
          If Val(Trim(grid1.text)) > 0 Then
             grid1.Col = 5
            If Val(Trim(grid1.text)) > 0 Then
               RS.AddNew
               grid1.Col = 1
               RS!invoiceNo = Val(Me.I_NO.text)
               RS!invoiceDate = Me.i_dt.text
               RS!Genledger = Trim(Me.Genledger.text)
               
               RS!subledger = Trim(Me.customercode.text)
               
               RS!Bookcode = Trim(grid1.text)
               grid1.Col = 3
               RS!QUANTITY = Trim(grid1.text)
               grid1.Col = 5
               RS!rate = Trim(grid1.text)
               grid1.Col = 7
               RS!amount = Trim(grid1.text)
               LAMOUNT = Val(Trim(grid1.text))
               grid1.Col = 4
               RS!PRINTORDER = Trim(grid1.text)
               grid1.Col = 6
               RS!discount = Trim(grid1.text)
               grid1.Col = 8
               RS!netamount = LAMOUNT - Trim(grid1.text)
               LAMOUNT = 0
               RS!agentname = Trim(Me.cmbAgentName.text)
               RS!fyear = session
               RS!setupid = setupid
               
               If kk.State = 1 Then kk.close
               kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,Apply from KitQry where kitcode='" & RS!Bookcode & "'", con
               While kk.EOF = False
                 If kk!Apply = "y" Then
                  con.Execute "insert into CASHB_Free(INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
                  " values('" & Val(Me.I_NO.text) & "','" & Format(Me.i_dt.text, "MM/dd/yyyy") & "','" & Trim(Me.Genledger.text) & "','" & Trim(Me.customercode.text) & "','" & kk!Bookcode & "','" & (kk!qty * RS!QUANTITY) & "','" & kk!rate & "','" & Trim(Me.cmbAgentName.text) & "','" & setupid & "','" & session & "','" & txtMark & "')"
                 End If
                 kk.MoveNext
                Wend

               RS.update
            End If
         End If
     End If
  Next
  RS.close
  grid1.TopRow = 1
  RS.Open "select * from CASHC where  " & stringyear & " and invoiceno<=0 ", con, adOpenDynamic, adLockOptimistic
  '/******
  'Dim I, x As Integer
  
  
  
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
              If temprs.State = 1 Then temprs.close
              If Edit Then
                 temprs.Open "select * from CASHCTMP WHERE INVOICENO = " & Val(Me.I_NO.text) & " and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
                 If frmEndPartTrans.vs.text <> "" Then
                    temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.text) + "'"
                    RS!Genledger = temprs!Genledger & ""
                    RS!subledger = temprs!subledger & ""
                    RS!DebitorCredit = temprs!DebitorCredit & ""
                    RS!RYN = temprs!RYN & ""
                End If
                temprs.close
              Else
                 temprs.Open "select * from INVOICEEND where type='" & searchForm & "' and  " & stringyear & "", con, adOpenDynamic, adLockReadOnly, adCmdText
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
      
      
      con.Execute ("delete  from CASHCTmp where  " & stringyear & "  and  INVOICENO = " + Trim(I_NO.text))
      SAVED = True
  
  End If
  
    If Me.station.text <> "" Then
    
    
    s11 = ""
    ss11 = ""
    
    s11 = InStr(1, Me.station.text, " ")
    If s11 <> 0 Then
    ss11 = Trim(Mid(Me.station.text, 1, s11))
    Else
    ss11 = Me.station.text
    End If
    PopUpValue1 = ss11

    
     UpdateDisPatchReg1 I_NO, i_dt, Me.customercode, PopUpValue1, Trim(Me.bundles), Trim(Me.cmbtransportname.text), "-", Trim(Me.biltno.text), Me.bdated, Trim(Me.freight), "CashRegister"
     PopUpValue1 = ""
    End If
    'End If

  
  If SAVED Then
      Unload frmEndPartTrans
   
      MsgBox "Record Saved"
      
      Me.customercode.Enabled = False
      Me.grid1.Enabled = False
      Me.Commandall.Enabled = False
      Me.Commandother.Enabled = False
      Me.Commandadd.Enabled = True
      Me.Commandedit.Enabled = True
      Me.Commandsearch.Enabled = True
      Me.Commandsave.Enabled = False
      Me.Commanddelete.Enabled = True
      Me.Commandabandon.Enabled = True
      Me.CommandPrint.Enabled = True
      Me.Commandprintnh.Enabled = True
   End If

addmode = False
addoredit = False
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
Me.Commandsave.Enabled = False


    If (AuditTrail = "y") Then
     
     If (txtchecked.text = "y") Then
     
         actionType_ = "Edit"
         vtype1_ = "CM"
         vtypeNew = "CM"
         vdate_ = Trim(i_dt.text)
         vno_ = Trim(I_NO.text)
         
         frmAuditTrailLog_Rem.Show 1
         
      End If
     
     End If


Exit Sub
save_:
MsgBox "" & err.DESCRIPTION

End Sub

Private Sub Commandsave_GotFocus()
txtAmtwords = toword(countersale.mna)
End Sub

Private Sub Commandsearch_Click()
   
   sqlQry = "select InvoiceNo,InvoiceDate,Subledger,NetAmount from CASHA where InvoiceNo"
   orderby = "order by InvoiceNo"

   
   searchType = "inv"
   popuplist10 "select InvoiceNo,InvoiceDate,Subledger,NetAmount from CASHA where " & stringyear & "  order by InvoiceNo", con

End Sub

Private Sub Commandsearch_GotFocus()
If PopUpValue1 <> "" Then
     I_NO.text = PopUpValue1
     I_NO_LostFocus
     
     If I_NO.Enabled = True Then
        I_NO.SetFocus
     End If
     
     SetButton Commandadd, Commandedit, Commandsave, Commanddelete
     PopUpValue1 = ""
End If

End Sub

Private Sub customercode_KeyPress(KeyAscii As Integer)
  ' If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
  '      SendKeys "{tab}"
  '      Exit Sub
  ' End If
    
   ' If KeyAscii = 13 Then
    '   SendKeys "{DOWN}"
    '   SendKeys "{TAB}"
   ' marka.SetFocus
    'End If
End Sub
Private Sub customercode_LostFocus()
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open "select * from sledger where gledger='SUNDRY DEBTORS' and subledger='" + Trim(customercode.text) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.RecordCount <= 0 Then
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
             Combosldistrictcode.text = RS!distcode
        Else
        Combosldistrictcode.text = RS!distcode
        End If
        Me.textbox.text = Me.customercode.text
        Me.customercode.Visible = False

End Sub

Private Sub Delete_Click()
If grid1.Row >= 1 Then
    grid1.SetFocus
    grid1.RemoveItem (grid1.Row)
    If grid1.Row > 1 Then
        grid1.Row = grid1.Row - 1
    End If
    Grid1_Click
End If
End Sub

Private Sub Form_Activate()
    
    mna.Enabled = True
    Label2.Enabled = True
    panel.Enabled = True
    
    Commandsave.Enabled = False
    Commanddelete.Enabled = False
    
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
    'txtMark.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If grid1.Row >= 1 Then
           grid1.RemoveItem grid1.Row
           a = grid1.text
           tempmeb.text = a
           a = templost
           grid1.SetFocus
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
          If addmode = True Then
                sendkeys "{DOWN}"
           End If
            sendkeys "{TAB}"
        Else
            If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("grid1")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("tempmeb")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bookname")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bundles")) Then
                sendkeys ("{TAB}")
            End If
        End If
    End If
End Sub


Private Sub Form_Load()

On Error Resume Next

Screen.MousePointer = vbHourglass

Me.top = 100
Me.Left = 100

Me.Width = 12720
Me.Height = 8985
grid1.Left = 150

Me.Caption = "Counter Sales"
      
    
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
    Me.top = 0
    Me.Left = 0
      
    grid1.rows = 2
    grid1.Cols = 1
    grid1.rows = 10
    grid1.Cols = 9
    grid1.Row = 0
    grid1.Col = 1
    grid1.text = "Book Code "
    grid1.Col = grid1.Col + 1
    grid1.text = "Book Name"
    grid1.Col = grid1.Col + 1
    grid1.text = "Quantity"
    grid1.Col = grid1.Col + 1
    grid1.text = "Print. Ord."
    grid1.Col = grid1.Col + 1
    grid1.text = "Rate"
    grid1.Col = grid1.Col + 1
    grid1.text = "Disc %"
    grid1.Col = grid1.Col + 1
    grid1.text = "Amount"
    grid1.Col = grid1.Col + 1
    grid1.text = "Disc. Amount"
    grid1.RowHeight(0) = grid1.CellHeight + 50
    grid1.ColWidth(0) = 150
    grid1.ColWidth(1) = 1100
    grid1.ColWidth(2) = 3250
    grid1.ColWidth(3) = 850
    grid1.ColWidth(4) = 850
    grid1.ColWidth(5) = 1200
    grid1.ColWidth(6) = 1000
    grid1.ColWidth(7) = 1400
    grid1.ColWidth(8) = 1400
    Me.CommandPrint.Enabled = True
    Me.Commandprintnh.Enabled = True
    
    
    '=================================================================================================
   ' Set RS = CON.Execute("exec BookQry '" & session & "'," & main.setupid & "")
    If RS.State = 1 Then RS.close
    RS.Open "select * from books where " & stringyear, CCON, adOpenDynamic, adLockReadOnly, adCmdText
    
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
    RS.Open "select subledger from sledger where gledger='" + Trim(Genledger.text) + "' and " & stringyear, CCON, adOpenDynamic, adLockReadOnly, adCmdText
    
    If Not RS.BOF Then
        Do While Not RS.EOF
            Me.customercode.AddItem RS("subledger")
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close

    
    
    '=====================================================================================================
    If RS.State = 1 Then RS.close
    
    RS.Open "select Distinct categorycode from DISCCATS where " & stringyear & " order by categorycode", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.cmbdiscountcat.AddItem RS!categorycode
            Me.cboCatII.AddItem RS!categorycode
            Me.cboCatII1.AddItem RS!categorycode
            
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    
    RS.close
    RS.Open "select  transportname from transportMaster where " & stringyear & " order by transportname", con, adOpenDynamic, adLockReadOnly, adCmdText
    cmbtransportname.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cmbtransportname.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    RS.close

  
 
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
    
      
    txtMark.ListIndex = 0
    

    
    
    Bookcode.Left = grid1.Left
    Bookcode.Visible = False
    Bookname.Visible = False
    grid1.rows = 100
    For I = 1 To 99
        grid1.RowHeight(I) = 300
    Next
    Bookcode.Width = 1230
    Bookname.Width = 2830
    amount.Width = rate.Width
    
    '===================================================================
    Me.cmbareaname.Clear
    Me.Combosldistrictcode.Clear
     If RS.State = 1 Then RS.close
    'RS.Open "select district District,[State],DistrictId from districtView", CON_blue
    Set RS = CON_blue.Execute("exec DistrictList")
    
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS!District
            'Me.cmbareaname.AddItem RS!DISTRICTNAME
            'Me.Combosldistrictcode.ItemData(Combosldistrictcode.NewIndex) = RS!DistrictID
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If

    
'    If RS.State = 1 Then RS.close
'    RS.Open "select * from DISTRICTS where " & stringyear & " order by DISTRICTNAME", con, adOpenDynamic, adLockReadOnly, adCmdText
'    If Not RS.EOF Then
'        Do While Not RS.EOF
'            Me.cmbareaname.AddItem RS!DISTRICTNAME
'            If Not RS.EOF Then
'                RS.MoveNext
'            End If
'        Loop
'    End If
'    RS.close
    
    
    '===================================================================
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT MAX(INVOICENO) FROM CASHA where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not IsNull(RS(0)) Then
       Me.I_NO.text = RS(0)
       countersale.Enabled = True
       countersale.Edit = False
       countersale.I_NO_LostFocus
       countersale.I_NO.Enabled = False
       lastrow = 0
       lastcol = 1
       
       Dim ctl As Control
       For Each ctl In countersale.Controls
           If Not TypeOf ctl Is CommandButton Then
              ctl.Enabled = False
           End If
           If UCase(Trim(ctl.Name)) = UCase(Trim(countersale.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(countersale.Commandall.Name)) Then
              ctl.Enabled = False
           End If
       Next
       countersale.Picture5.Enabled = True
       addoredit = False
       sendkeys "{TAB}"
   Else
       'kk.Open "SELECT MAX(INVOICENO) FROM CASHA", CON, adOpenDynamic, adLockReadOnly, adCmdText
       'If kk(0) <> "" Then
       '   Me.I_NO.Text = Trim(str(kk(0) + 1))
       'Else
       '   Me.I_NO.Text = "1"
       'End If
       'kk.Close
   End If
   RS.close
   
'    If kk.State = 1 Then kk.close
'    kk.Open "SELECT MAX(INVOICENO) FROM CASHA where " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
'    If kk(0) <> "" Then
'        'Me.I_NO.Text = Trim(Str(kk(0) + 1))
'        Me.I_NO.Text = kk(0)
'        I_NO_LostFocus
'    Else
'       ' Me.I_NO.Text = "1"
'    End If
'    kk.close
   
   
   
   mna.Enabled = True
   Label2.Enabled = True
   Commanddelete.Enabled = True
   Commandedit.Enabled = True
   Commandsave.Enabled = False
   lastrow = 0
   lastcol = 1
   For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton Then
            ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    
    
''    Me.Combosldistrictcode.Clear
      Picture5.Enabled = True
''    If RS.State = 1 Then RS.close
''    RS.Open "select district District,[State],DistrictId from districtView", CON_blue
''
''    If Not RS.EOF Then
''        Do While Not RS.EOF
''            Me.Combosldistrictcode.AddItem RS!District
''            Me.Combosldistrictcode.ItemData(Combosldistrictcode.NewIndex) = RS!DistrictID
''            If Not RS.EOF Then
''                RS.MoveNext
''            End If
''        Loop
''    End If
''
''
''    If RS.State = 1 Then RS.close
''    RS.Open "select * from DISTRICTS where " & stringyear & " order by DISTRICTNAME", con, adOpenDynamic, adLockReadOnly, adCmdText
''    If Not RS.EOF Then
''        Do While Not RS.EOF
''            Me.cmbareaname.AddItem RS!DISTRICTNAME
''            If Not RS.EOF Then
''                RS.MoveNext
''            End If
''        Loop
''    End If
''    RS.close
    
    
    'SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
    
    
    BackColorFrom Me
    Commandsave.Enabled = False
    Commanddelete.Enabled = False

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub
Function returnCategory(s As String) As String
    Dim s1 As New ADODB.Recordset
    If s1.State = 1 Then s1.close
    
    s1.Open "select category from [groups] where groupcode='" & s & "' and " & stringyear, con
    If s1.EOF = False Then
       returnCategory = s1(0)
    End If
    
End Function


Sub Grid1_Click()
If Trim(Me.customercode.text) <> "" Then
Dim PREVROW As Integer
Dim prevcol As Integer
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
prevcol = Me.grid1.Col
PREVROW = Me.grid1.Row

If Me.grid1.Row > 1 Then
    grid1.Row = grid1.Row - 1
    grid1.Col = 1
    If Trim(grid1.text) <> "" Then
        grid1.Row = PREVROW
        grid1.Col = prevcol
        If Trim(Me.customercode.text) <> "" Then
            If Me.customercode.Enabled = True Then
                Me.customercode.Enabled = False
            End If
            grid1.Col = 1
            If prevcol > 1 And Trim(grid1.text) = "" Then
                grid1.Col = 2
                sendkeys Chr(13)
            Else
                grid1.Col = prevcol
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
        Me.grid1.Col = 1
        If prevcol > 1 And Trim(grid1.text) = "" Then
            Me.grid1.Col = 2
            Me.grid1.SetFocus
            sendkeys Chr(13)
        Else
        'IF GRID1.COL
            Me.grid1.Col = prevcol
            Me.grid1.SetFocus
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
            
            Select Case grid1.Col
            Case 1, 3, 4, 5, 6
                
                Bookname.Visible = False
                tempmeb.Visible = True: tempmeb.Enabled = True
                tempmeb.ZOrder
                
                If grid1.Col <> 1 Then
                    If grid1.Col <> 3 Then
                        tempmeb.text = Format(grid1.text, "0.00")
                        
                    Else
                        tempmeb.text = Format(grid1.text, "0")
                    End If
                   
                Else
                    tempmeb.text = grid1.text
                End If
                tempmeb.Width = grid1.ColWidth(grid1.Col)
                tempmeb.Left = grid1.CellLeft + leftAlign + 150
                tempmeb.top = grid1.top + grid1.CellTop '- 50
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.text = grid1.text
                Bookname.top = grid1.top + grid1.CellTop
                Bookname.Left = grid1.CellLeft + leftAlign
                Bookname.Width = grid1.ColWidth(grid1.Col)
            Case 6
                
            Case Else
                Bookname.Visible = False
                tempmeb.Visible = False
            End Select
            Select Case grid1.Col
                Case 1, 3, 4, 5, 6
                    tempmeb.Mask = ""
                    tempmeb.MaxLength = 20
                Case 2
                    With Bookname
                        .Visible = True
                        .ZOrder
                    End With
                End Select
            Select Case grid1.Col
            Case 2
                Bookname.SetFocus
                If KeyAscii <> 13 Then
                    sendkeys Chr(KeyAscii)
                End If
            Case 1, 3, 4, 5, 6
                mprevcol = grid1.Col
                tempmeb.SetFocus
            Case Else
                If KeyAscii = 13 Then
                    sendkeys "{RIGHT}"
                End If
            End Select
        End If
    If maxrow < grid1.Row Then
        maxrow = grid1.Row
    End If
End If
    lastrow = grid1.Row
    lastcol = grid1.Col
    
End If
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu dd, , grid1.Left + X, grid1.top + Y
End If
End Sub

Private Sub Grid1_Scroll()
If tempmeb.Visible <> True And Bookname.Visible <> True Then
        Bookname.Visible = False
        tempmeb.Visible = False
        grid1.SetFocus
End If
autoscroll = True
End Sub
Private Sub i_dt_LostFocus()

If IsDate(i_dt) Then
  If checkData_ForThisNumber("casha", I_NO, i_dt) = True Then
      MsgBox "Please select valid Invoice No. for this date.."
      i_dt.SetFocus
  End If
End If

'''On Error Resume Next
'''If Trim(i_dt.Text) <> Trim("__/__/____") Then
'''    If Not checkdate(Trim(i_dt.Text), i_dt) Then
'''        i_dt.SetFocus
'''    End If
'''    Dim tRS1 As New ADODB.Recordset
'''    Dim trs2 As New ADODB.Recordset
'''
'''     If trs2.State = 1 Then trs2.close
'''    trs2.Open "Select invoiceno as cn from casha where " & stringyear, con, adOpenDynamic, adLockOptimistic
'''    If trs2.RecordCount <= 0 Then
'''       Exit Sub
'''    Else
'''        If tRS1.State = 1 Then tRS1.close
'''        tRS1.Open "Select min(invoiceno) as mid,invoicedate from casha where " & stringyear & " group by invoiceno,invoiceDate", con, adOpenDynamic, adLockOptimistic
'''        If tRS1.RecordCount > 0 Then
'''
'''            If CDate(i_dt) <= tRS1!INVOICEDATE Then
'''
'''               If CDate(i_dt) <> tRS1!INVOICEDATE Then
'''               If Month(CDate(i_dt)) <> 4 And Day(CDate(i_dt)) <> 1 Then
'''                 MsgBox "Please select valid Cash Memo No. for this date.."
'''                 i_dt.SetFocus
'''                 Exit Sub
'''
'''            Else
'''                If tRS1!Mid <> 1 Then
'''               If Val(I_NO) >= tRS1!Mid Then
'''                 MsgBox "Please select Cash Memo No. for this date.."
'''                 i_dt.SetFocus
'''                 Exit Sub
'''               End If
'''               End If
'''                 End If
'''
'''               End If
'''            End If
'''        End If
'''    End If
'''
'''
'''
'''    If trs2.State = 1 Then trs2.close
'''    trs2.Open "Select max(invoiceno) as mid from casha where " & stringyear & " and invoicedate <= cdate('" & i_dt.Text & "')-1", con, adOpenDynamic, adLockOptimistic
'''    If trs2.RecordCount > 0 Then
'''        If IsNull(trs2!Mid) <> True Then
'''            If Val(I_NO.Text) >= trs2!Mid Then
'''               If tRS1.State = 1 Then tRS1.close
'''               tRS1.Open "Select  min(InvoiceNo)as m2 from casha where " & stringyear & "  and invoicedate >= cdate('" & i_dt.Text & "')+1", con, adOpenDynamic, adLockOptimistic
'''               If tRS1.RecordCount > 0 Then
'''                  If IsNull(tRS1!m2) <> True Then
'''                     If Val(I_NO.Text) <= tRS1!m2 Then
'''
'''                     Else
'''                         MsgBox "Please select valid Cash Memo No. for this date.."
'''                         I_NO.SetFocus
'''                     End If
'''                  End If
'''               End If
'''
'''            Else
'''               If I_NO.Enabled = False Then Exit Sub
'''                    If i_dt.Enabled = True Then
'''                            MsgBox "Please select valid Cash Memo No. for this date.."
'''                            I_NO.Enabled = True
'''                            I_NO.SetFocus
'''                    End If
'''                 End If
'''            End If
'''     End If
'''Else
'''    i_dt.SetFocus
'''    HIT
'''    Exit Sub
'''End If
End Sub


Private Sub I_DTOB_LostFocus()
If Trim(I_DTOB.text) <> "__/__/____" Then
    If Not checkdate(Trim(I_DTOB.text), I_DTOB) Then
        I_DTOB.SetFocus
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
Sub I_NO_LostFocus()
    On Error Resume Next
    
    If Val(inviceNo) > 0 Then
       I_NO.text = inviceNo
    End If
    
    inviceNo = ""
    
    
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    If Trim(I_NO.text) = "" Then
        MsgBox "Cash Memo No cannot be null"
        I_NO.SetFocus
    Else
        If RS.State = 1 Then RS.close
        RS.Open "Select * from  CASHA where INVOICENO = " + Trim(I_NO.text) + " and " & stringyear, con, adOpenStatic, adLockReadOnly
        If RS.EOF Then
            If addoredit = False Then
            '     MsgBox "Cash Memo No not found"
            '     Exit Sub
            End If
            Exit Sub
        End If
        
        If addoredit Then
            MsgBox "Cash Memo No already exist..."
            'I_NO.SetFocus
            HIT
            Exit Sub
        End If
        
        
        invoiceabandon
        
''        Dim ctl As Control
''        For Each ctl In Me.Controls
''           If Not TypeOf ctl Is CommandButton Then
''                ctl.Enabled = True
''            End If
''        Next
        
        If (AuditTrail = "y") Then
         
         If (RS!Checked_YesNo = True) Then
           txtchecked.text = "y"
        Else
            txtchecked.text = "n"
        End If
        
        End If
        
        
        
        Me.Commandother.Enabled = True
        I_NO.text = RS!invoiceNo
        
        txtSchool = RS!scname & ""
        txtScId.text = RS!scid & ""
        lblState.Caption = RS!states & ""

        txtAmtwords.text = RS!Amtwords & ""
        Me.i_dt.text = RS!invoiceDate
        Me.Genledger.text = Trim(RS!Genledger)
        Me.customercode.text = Trim(RS!subledger)
        Me.textbox.text = Trim(RS!subledger)
        Me.I_OB.text = Trim(RS!orderby)
        Me.I_DTOB.text = IIf(IsNull(RS!ORDERDATE), "__/__/____", RS!ORDERDATE)
        'Me.marka.Text = Trim(rs!marka)
        Me.bundles = Trim(RS!bundles)
        txtMark = RS!Godown & ""
        'Me.through.Text = rs!through
        'Me.through1.Text = rs!through1
        Me.station.text = RS!station
        Me.biltno.text = Trim(RS!biltyno)
        Me.bdated = IIf(IsNull(RS!BILTYDATE), "__/__/____", RS!BILTYDATE)
        Me.freight = Trim(RS!freight)
        'Me.weight = Trim(rs!weight)
        Me.labelbybank = Format(Round(Val(RS!baa), 2), "0.00")
        mna.Caption = Format(Round(Val(RS!netamount), 2), "0.00")
        Me.cmbtransportname.text = IIf(IsNull(RS!transportname), "", RS!transportname)

      
        
        If RS!District <> "" Then
            Combosldistrictcode.text = RS!District
        End If
        
        lblDId.Caption = RS!cityId & ""
        
        If Me.customercode.text = "CASH PARTY" Then
            textbox.text = RS!CASHPARTYNAME
            Optioncash = True
            Me.cmbAgentName.text = IIf(IsNull(RS!agentname), "", Trim(RS!agentname))
        Else
            Optioncredit = True
            Me.cmbAgentName.text = IIf(IsNull(RS!agentname), "", Trim(RS!agentname))
        End If
        cmbdiscountcat.text = IIf(IsNull(RS!discat), "", RS!discat)
        cboCatII.text = IIf(IsNull(RS!discatII), "", RS!discatII)
        cboCatII1.text = IIf(IsNull(RS!discatIII), "", RS!discatIII)
        
        RS.close
        grid1.TopRow = 1
    '*/**/*/*/*/*//*/*
    If RS.State = 1 Then RS.close
    RS.Open "Select * from CASHB where INVOICENO =" + Trim(I_NO.text) + " and " & stringyear & "  order by SNO ", con, adOpenStatic, adLockReadOnly
    If Not RS.EOF Then
            grid1.Row = 1
            grid1.Col = 1
            Do While Not RS.EOF
               If Trim(RS!invoiceNo) = Trim(I_NO.text) Then
                grid1.Col = 1
                grid1.text = Trim(RS!Bookcode)
                If kk.State = 1 Then
                    kk.close
                End If
                kk.Open "select * from books where bookcode='" + Trim(RS!Bookcode) + "' and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
                grid1.Col = 2
                grid1.text = Trim(kk!Bookname)
                grid1.Col = 3
                grid1.text = Trim(RS!QUANTITY)
                grid1.Col = 5
                grid1.text = Format(Round(RS!rate, 2), "0.00")
                grid1.Col = 7
                grid1.text = Format(Round(RS!amount, 2), "0.00")
                grid1.Col = 4
                grid1.text = Format(Round(RS!PRINTORDER, 2), "0.00")
                grid1.Col = 6
                grid1.text = Format(Round(RS!discount, 2), "0.00")
                grid1.Col = 8
                grid1.text = Format(Round(RS!amount * (RS!discount / 100), 2), "0.00")
                End If
                If Not RS.EOF Then
                    RS.MoveNext
                End If
                grid1.Row = grid1.Row + 1
                grid1.rows = grid1.rows + 1
            Loop
            maxrow = grid1.Row
        End If
        Row = grid1.Row
        Col = grid1.Col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            grid1.Row = I
            grid1.Col = 7
            totalamount = totalamount + Val(Trim(grid1.text))
            grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.text))
        Next
        mga.Caption = Format(Round(totalamount, 2), "0.00")
        mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
        Me.tqu.Caption = ""
        For I = 1 To maxrow
            grid1.Col = 3
            grid1.Row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(grid1.text))
        Next
        grid1.Row = RRR
        grid1.Col = CCC
    End If
    mna.Enabled = True
    Label2.Enabled = True
    Me.Commandother.Enabled = True
    
    Commandsave.Enabled = False
    Commanddelete.Enabled = False
    
End Sub
Private Sub I_OB_GotFocus()
Dim trs As New ADODB.Recordset
trs.Open " SELECT DISTCODE    FROM SLEDGER  WHERE " & stringyear & " and  SUBLEDGER='" & customercode.text & "' ", con, adOpenStatic, adLockOptimistic, adCmdText
       If Not trs.BOF Then
           If Combosldistrictcode.text = "" Then
               Combosldistrictcode.text = IIf(IsNull(trs!distcode), "", trs!distcode)
          End If
      End If

End Sub

Private Sub I_OB_LostFocus()
I_OB = UCase(I_OB)
End Sub

Private Sub marka_LostFocus()
marka = UCase(marka)
End Sub

Private Sub Optioncash_Click()
If Optioncash.value = True Then
       
''       Label4.Visible = True
''       Combosldistrictcode.Visible = True
''       Combosldistrictcode.Enabled = True
''       cmbdiscountcat.Visible = True
''       lbldis(0).Visible = True
''       lbldis(1).Visible = True
''       cboCatII.Visible = True
''       cboCatII1.Visible = True
''       lbldis(2).Visible = True

       
 End If


End Sub

Private Sub Optioncredit_Click()
If Optioncredit.value = True Then
       Label4.Visible = False
       Combosldistrictcode.Visible = True
       'lbldis(0).Visible = False
       'lbldis(1).Visible = False
       cmbdiscountcat.Visible = False
       
       cboCatII.Visible = False
       cboCatII1.Visible = False
       'lbldis(2).Visible = False


End If

End Sub

Private Sub station_LostFocus()
station = UCase(station)
End Sub

Private Sub tempmeb_Change()
If grid1.Col = 1 Or grid1.Col = 2 Then
    grid1.text = tempmeb.text
Else
    If grid1.Col = 3 Then
        grid1.text = Format(tempmeb.text, "0")
    Else
        grid1.text = Format(tempmeb.text, "0.00")
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
            Select Case grid1.Col
                Case 1
                    'If RS.State = 1 Then
                    '    RS.close
                    'End If
                    'RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                    Set RS = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(grid1.text) & "'")
                    If Not RS.BOF Then
                        'RS.MoveFirst
                        'RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF And Trim(grid1.text) <> "" Then
                            RS.close
                            Exit Sub
                        Else
                            RS.close
                        If Trim(grid1.text) <> "" Then
                                grid1.Col = 3
                            Else
                                grid1.Col = 2
                            End If
                        End If
                    Else
                        If Trim(grid1.text) <> "" Then
                            grid1.Col = 3
                        Else
                            grid1.Col = 2
                        End If
                    End If
                    grid1.SetFocus
                    Grid1_Click
                Case 3
                    If Val(tempmeb.text) > 0 Then
                        'Grid1.Col = Grid1.Col + 2
                        grid1.Col = 1
                        grid1.Row = grid1.Row + 1
                        grid1.rows = grid1.rows + 1
                        
                        grid1.SetFocus
                        Grid1_Click
                    End If
                Case 4
                    grid1.Col = grid1.Col + 2
              '       SendKeys "{LEFT}"
                    grid1.SetFocus
                    Grid1_Click
                Case 5
                    If Val(tempmeb.text) > 0 Then
                        grid1.Col = grid1.Col - 1
                        grid1.SetFocus
                        Grid1_Click
                    End If
                Case 6
                    If Val(grid1.TextMatrix(grid1.Row, 4)) <> Val(grid1.TextMatrix(grid1.Row, 6)) Then
                      MsgBox "Discount And Printorder  Not Match.."
                      
                   End If
                    grid1.Col = 1
                    grid1.Row = grid1.Row + 1
                    grid1.rows = grid1.rows + 1
                    grid1.SetFocus
                    Grid1_Click
            End Select
        Else
        If grid1.Col = 3 Or grid1.Col = 4 Or grid1.Col = 5 Or grid1.Col = 6 Then
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
If Optioncash = False Then
    Me.customercode.Enabled = True
    Me.customercode.Visible = True
  '  Me.customercode.Height = 1100
    Me.customercode.ZOrder
    Me.customercode.SetFocus
   

End If

End Sub

Private Sub textbox_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Then
   cmbareaname.Visible = True
   Me.customercode.Enabled = True
   Me.customercode.Visible = False
  'Me.customercode.Height = 1100
   Me.cmbareaname.ZOrder
   Me.cmbareaname.SetFocus
   

Else
   cmbareaname.Visible = False
End If
End Sub

Private Sub textbox_LostFocus()

 If Optioncash = True And textbox.text <> "" Then
        Me.customercode.text = "CASH PARTY"
 End If
 textbox.text = UCase(textbox.text)
 
End Sub

Private Sub through_LostFocus()
through = UCase(through)
End Sub

Private Sub through1_LostFocus()
through1 = UCase(through1)
End Sub
Private Sub weight_KeyPress(KeyAscii As Integer)
  
End Sub

Private Sub weight_LostFocus()
weight = UCase(weight)

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
      If flagyes = True Then
          CNSetup
          kkk.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
          If Not kkk.BOF Then
             Print #1, Chr(27) + Chr(15) + Chr(14)
             Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname)
             Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
             Print #1, Tab((paperWidth - Len(Trim(kkk!phone1))) / 2); Trim(kkk!phone1)
             Line = Line + 3
         End If
          If rs1.State = 1 Then
             rs1.close
          End If
          Print #1, Chr(27) + Chr(14)
          Line = Line + 1
      
    Else
       Print #1, ""
       Print #1, ""
       Print #1, ""
       Print #1, Chr(27) + Chr(15)
       Line = Line + 4
  End If
  
  
  
    If rs1.State = 1 Then
        rs1.close
    End If
    rs1.Open "select top 100 * from CASHA where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.text) + "'"
   
    Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!subledger; Tab(T5); "Cash Memo No. : "; Trim(rs1!invoiceNo); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!invoiceDate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.close
            End If
            kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE
                Print #1, Tab(3); kkk!address1; Tab(T5); "Order by : "; Trim(rs1!orderby); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!ORDERDATE
                Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.: "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!BILTYDATE
                kkk.close
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
                kk.close
            End If
            kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!invoiceNo) + " and " & stringyear & " order by discount,printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                        tdata.Open "select sum(amount) from CASHB where invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " and " & stringyear & " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
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
           Print #1, Tab(T5 - 6); rsets(Trim(Str(totalquantity)), 7); Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
           Print #1, Tab(T6); repli("-", 22)
           If kk.State = 1 Then
                kk.close
           End If
           kk.Open "Select * from CASHC where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
           kk.Open "Select * from CASHA where  " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
            tempdata.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdTable
            Print #1, Tab(1); "E.& O.E"
            Print #1, Tab(LEFTM); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!cname)) * 0.75))); "FOR " + Trim(tempdata!cname)
            Print #1, Tab(LEFTM); ""

       'PRINT THE FOOTER IN INVOICE END
       
        
        
        Close #1
        PrintOption.Show


End Sub
'****************************56


Sub bakupprintinvoice()
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
MaxLine = 66
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
If Printheader = True Then
CNSetup
kkk.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
If Not kkk.BOF Then
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, Chr(27) + Chr(71); Chr(27) + Chr(15) + Chr(14)
      Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname)
      Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
      Print #1, Tab((paperWidth - Len(Trim(kkk!phone1))) / 2); Trim(kkk!phone1)
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
If Printheader = True Then
   Print #1, Tab(T7 + 4); IIf(Printheader = True, kkk!uptt, "")
   Print #1, Tab(T7 + 4); kkk!cst; Chr(27) + Chr(72)
   Line = Line + 2
Else
   Print #1, Tab(0); Chr(27) + Chr(15);
End If

Print #1, Tab(0); repli("*", 148)
Print #1, Tab(0); Chr(27) + Chr(15) + Chr(14); Tab(30); "CASH MEMO"; Chr(20)
Print #1, Tab(0); repli("*", 148)
Line = Line + 3
If rs1.State = 1 Then
   rs1.close
End If

If rs1.State = 1 Then
    rs1.close
End If
rs1.Open "select * from CASHA where " & stringyear & " and invoiceno='" + Trim(Me.I_NO.text) + "'", con, adOpenDynamic, adLockReadOnly
'rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To,"; Tab(T1 - 3); IIf(Optioncash.value = True, rs1!CASHPARTYNAME, rs1!subledger); Tab(T5); "Cash Memo No. : "; Trim(rs1!invoiceNo); Tab(T8 + 5); "Dt. :"; Tab(T8 + 12); rs1!invoiceDate 'Chr(27) + Chr(15);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); kkk!DESCFORINVOICE
        Print #1, Tab(3); kkk!address1; Tab(T5); "Order by     : "; Trim(rs1!orderby); Tab(T8 + 5); "Dt. :"; Tab(T8 + 12); rs1!ORDERDATE
        Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.    : "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. :"; Tab(T8 + 12); rs1!BILTYDATE
        kkk.close
        Print #1, "Through  :"; Tab(12); Trim(rs1!through) + ", " + Trim(rs1!through1)
        Print #1, "Station  :"; Tab(12); Trim(rs1!station); Tab(T5); "Pvt. Mark    : "; Trim(rs1!marka)
        Print #1, "Freight  :"; Tab(12); Trim(rs1!freight); Tab(T5); "Weight       : "; Trim(rs1!weight); Tab(T7 + 7); "Bundle(s)   : "; Trim(rs1!bundles); Chr(27) + Chr(72)
        Print #1, repli("-", 150)
        Print #1, "S.No."; Tab(11); "Book Description"; Tab(T5 - 3); "Quantity"; Tab(T6 + 4); "Rate"; Tab(T7 + 4); "Amount"; Tab(T8 + 9); "Net Amount"
        Print #1, repli("-", 150)
        Line = Line + 9
        
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
    kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!invoiceNo) + " and " & stringyear & " order by discount,printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "select sum(amount) from CASHB where invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " and " & stringyear & " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                    Print #1, Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0), 2)), "0.00")), 12)
                    Print #1, Tab(T5); "Less Discount @ " + Trim(Format(Str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(Str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(Str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                    Line = Line + 2
                    netamount = netamount + Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                End If
                tdata.close
                Print #1, Tab(T7); repli("-", 22)
                Line = Line + 1
                Loop
            End If
        End If
        Print #1, Tab(T5 - 10); repli("-", 22)
        Print #1, Tab(T5 - 4); rsets(Trim(Str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
        
        Line = Line + 2
        
        
        If kk.State = 1 Then
             kk.close
        End If
        kk.Open "Select * from CASHC where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
        f1 = 1
        If Not kk.BOF Then
            
            Do While Not kk.EOF
                If kk!amount > 0 Then
                   If f1 = 1 Then
                      Print #1, Tab(T8); repli("-", 22)
                      Line = Line + 1
                      f1 = 2
                    End If
                    f1 = 2
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
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T8); repli("-", 22)
            Print #1, Chr(27) + Chr(71); Tab(T5); "NET AMOUNT: "; Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
            VNetamt = netamount
            Line = Line + 2
        End If
        kk.close
        kk.Open "Select * from CASHA where " & stringyear & " and  invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(T5); kk!txt1; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + Round(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(T5); kk!txt2; Tab(T8 + 5); rsets(Trim(Format(Str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + Round(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(T5); "CASH RECD. "; Tab(T8 + 4); rsets(Trim(Format(Str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - Round(kk!baa, 2)
             End If
        End If
        Print #1, Tab(T8); repli("-", 22)
        Print #1, Tab(T5); Chr(27) + Chr(71); "BALANCE : "; Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
        Line = Line + 2
        
        Do While Line < 65
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
        tempdata.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!cname)) * 0.65))); "FOR " + Trim(tempdata!cname)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Close #1
        
         PrintOption.Show

End Sub

Sub printinvoice111()
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
Dim FooterYes As Boolean
Dim totalquantity As Long
Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
Dim RS As ADODB.Recordset
Dim LEFTM As Integer
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
Set kkk = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Open "" + VB.App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
LEFTM = 5
FooterYes = False
header:
    If kkk.State = 1 Then
          kkk.close
    End If
    CNSetup
    kkk.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If FooterYes = True Then
        If Line > MaxLine - 5 Then
            Do While Line < 61
                Print #1, ""
                Line = Line + 1
            Loop
        End If
        Line = 0
        LEFTM = 5
        Print #1, Tab(0); repli("-", 96)
        
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); kkk!COURT; Tab(75); "FOR " + Trim(kkk!cname)
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
     Print #1, Chr(27) + Chr(77)
     Line = Line + 7
End If
Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CASH MEMO")))) / 2 - 10); Chr(14); "***CASH MEMO***"; Chr(20); Tab(50); IIf(Printheader = True, kkk!uptt, "")
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
rs1.Open "select top 100 * from CASHA where " & stringyear & " and invoiceno = " & Me.I_NO.text, con, adOpenDynamic, adLockReadOnly
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); " To,"; Tab(7); IIf(Optioncash.value = True, "", Mid$(rs1!subledger, 1, 5)); Tab(48); "Cash Memo No. : "; Trim(rs1!invoiceNo); Tab(82); "Dt. : "; rs1!invoiceDate; Chr(27) + Chr(72);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & IIf(Optioncash.value = True, rs1!CASHPARTYNAME, kkk!DESCFORINVOICE)
        Print #1, Tab(5); IIf(IsNull(kkk!address1), "", kkk!address1); Tab(49); Chr(27) + Chr(71); "Order by    : "; Chr(27) + Chr(72); Trim(rs1!orderby); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!address2), "", kkk!address2); Tab(49); Chr(27) + Chr(71); "Bilty No.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!address3), "", kkk!address3)
        kkk.close
        Print #1, Chr(27) + Chr(71); "Through  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!through) + IIf(Trim(rs1!through1) = "", "", "," & rs1!through1)
        Print #1, Chr(27) + Chr(71); "Station  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!station); Tab(71); Chr(27) + Chr(71); "Pvt. Mark   : "; Chr(27) + Chr(72); Trim(rs1!marka)
        Print #1, Chr(27) + Chr(71); "Freight  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!freight); Tab(40); Chr(27) + Chr(71); "Weight    : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."; Tab(73); Chr(27) + Chr(71); "Bundle(s)   : "; Chr(27) + Chr(72); Trim(rs1!bundles)
        Print #1, Chr(27) + Chr(71); repli("-", 96)
        Print #1, Tab(0); "S.No."; Tab(15); "Book Description"; Tab(50); "Quantity"; Tab(62); "Rate"; Tab(74); "Amount"; Tab(86); "Net Amount"
        Print #1, repli("-", 96); Chr(27) + Chr(72)
        Line = Line + 10
    End If
    If called1 Then
        GoTo printagain1
    End If
    If called2 Then
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.close
    kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!invoiceNo) + " and " & stringyear & " order by discount,sno ", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                Print #1, Tab(0); rsets(Trim(Str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(Str(kk!QUANTITY)), 5); Tab(58); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!QUANTITY
                Line = Line + 1
                If Line > MaxLine - 4 Then
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
            If Line > MaxLine - 4 Then
                    called2 = True
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
                    
                    
printagain2:
                    called2 = False
                End If
                Print #1, Tab(70); repli("-", 12)
                Line = Line + 1
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CashB where invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " and " & stringyear & " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(Str(tdata(0)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(70); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.close
             Loop
         End If
    End If
    Print #1, repli("-", 96)
    Print #1, Tab(52); rsets(Trim(Str(totalquantity)), 5); Tab(84); rsets(Trim(Format(Str(netamount), "0.00")), 12)
    Line = Line + 2
    If kk.State = 1 Then kk.close
    kk.Open "Select * from CASHC where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
       Do While Not kk.EOF
             If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(60); Trim(kk!text) + "    " + Trim(Format(Str(kk!rate), "0.00")); Tab(84); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(60); Trim(kk!text); Tab(84); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
          
        End If
        Print #1, Tab(84); repli("-", 12)
        Print #1, Chr(27) + Chr(71); Tab(60); "NET AMOUNT: "; Tab(85); rsets(Trim(Format(Str(netamount), "0.00")), 12); Chr(27) + Chr(72)
        VNetamt = netamount
        Line = Line + 2
        kk.close
        kk.Open "Select * from CASHA where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.text), con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(60); kk!txt1 & "    :"; Tab(84); rsets(Trim(Format(Str(Abs(kk!txt1a)), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + kk!txt1a
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(60); kk!txt2 & " :"; Tab(84); rsets(Trim(Format(Str(Abs(kk!txt2a)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + kk!txt2a
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(59); "CASH RECD. :"; Tab(84); rsets(Trim(Format(Str(Abs(kk!baa)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - kk!baa
             End If
             If netamount <> 0 Then
                 Print #1, Tab(84); repli("-", 12)
                 Print #1, Tab(59); Chr(27) + Chr(71); "BALANCE   : "; Tab(85); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
                 Print #1, Tab(84); repli("-", 12);
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
        tempdata.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdTable
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
        Print #1, ""
        'PRINT THE FOOTER IN INVOICE END
        Close #1
        PrintOption.Show
        
End Sub
Private Sub txtschool_GotFocus()
If PopUpValue1 <> "" Then
   
txtScId = PopUpValue1
txtSchool.text = PopUpValue2 & ", " & PopUpValue3
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""

End If
End Sub

Private Sub txtschool_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   
   Screen.MousePointer = vbHourglass
   tblNo = 9
   frmSearchItem.Show
   Screen.MousePointer = vbDefault
   
End If

End Sub
