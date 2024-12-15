VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmInvoice_blueprint 
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   10170
   Begin VB.Frame panel 
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7770
      Left            =   60
      TabIndex        =   19
      Top             =   0
      Width           =   10020
      Begin VB.ComboBox customercode 
         Appearance      =   0  'Flat
         Height          =   960
         Left            =   5880
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Top             =   420
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox Genledger 
         Height          =   315
         Left            =   6255
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   6450
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00B8E4F1&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   195
         ScaleHeight     =   690
         ScaleWidth      =   9510
         TabIndex        =   25
         Top             =   6885
         Width           =   9510
         Begin VB.CommandButton Commandhelp 
            Caption         =   "Help"
            Height          =   495
            Left            =   -720
            TabIndex        =   35
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&dit"
            Height          =   585
            Left            =   1072
            Picture         =   "frmInvoice_blueprint.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Height          =   585
            Left            =   2129
            Picture         =   "frmInvoice_blueprint.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   585
            Left            =   3186
            Picture         =   "frmInvoice_blueprint.frx":1026
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Enabled         =   0   'False
            Height          =   585
            Left            =   4243
            Picture         =   "frmInvoice_blueprint.frx":15B0
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   585
            Left            =   5300
            Picture         =   "frmInvoice_blueprint.frx":2194
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   60
            Width           =   975
         End
         Begin VB.CommandButton CommandPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Enabled         =   0   'False
            Height          =   585
            Left            =   7414
            Picture         =   "frmInvoice_blueprint.frx":2D78
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   585
            Left            =   8475
            Picture         =   "frmInvoice_blueprint.frx":395C
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandadd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   585
            Left            =   15
            Picture         =   "frmInvoice_blueprint.frx":4540
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton Commandprintnh 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N&HPrint"
            Enabled         =   0   'False
            Height          =   585
            Left            =   6357
            Picture         =   "frmInvoice_blueprint.frx":5124
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   45
            Width           =   975
         End
      End
      Begin VB.ComboBox Bookcode 
         Height          =   2325
         ItemData        =   "frmInvoice_blueprint.frx":5D08
         Left            =   2700
         List            =   "frmInvoice_blueprint.frx":5D0A
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   24
         Top             =   2850
         Width           =   2355
      End
      Begin VB.ComboBox Bookname 
         Height          =   960
         Left            =   2700
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   23
         Top             =   2730
         Width           =   2355
      End
      Begin VB.CommandButton Commandother 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&End Part"
         Enabled         =   0   'False
         Height          =   480
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6150
         Width           =   930
      End
      Begin VB.CommandButton Commandall 
         Caption         =   "All Books"
         Height          =   420
         Left            =   -315
         TabIndex        =   21
         Top             =   6075
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ComboBox cmbAgentName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5895
         TabIndex        =   6
         Top             =   825
         Width           =   3840
      End
      Begin VB.TextBox txtadst 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2205
         TabIndex        =   20
         Top             =   6450
         Width           =   1035
      End
      Begin VB.ComboBox cmbtransportname 
         Height          =   315
         Left            =   1410
         TabIndex        =   13
         Top             =   2010
         Width           =   1920
      End
      Begin VB.ComboBox txtMark 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmInvoice_blueprint.frx":5D0C
         Left            =   8250
         List            =   "frmInvoice_blueprint.frx":5D19
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1380
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3705
         Left            =   150
         TabIndex        =   18
         Top             =   2400
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   6535
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
         Left            =   5910
         TabIndex        =   4
         Top             =   450
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox through 
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   1410
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_DTOB 
         Height          =   315
         Left            =   3900
         TabIndex        =   3
         Top             =   780
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox bundles 
         Height          =   285
         Left            =   1155
         TabIndex        =   8
         Top             =   1410
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_OB 
         Height          =   315
         Left            =   2250
         TabIndex        =   2
         Top             =   780
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
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
         Left            =   1230
         TabIndex        =   1
         Top             =   780
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tempmeb 
         Height          =   285
         Left            =   210
         TabIndex        =   37
         Top             =   2490
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rate 
         Height          =   285
         Left            =   1110
         TabIndex        =   38
         Top             =   4530
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox amount 
         Height          =   285
         Left            =   150
         TabIndex        =   39
         Top             =   4920
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_NO 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Top             =   780
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox through1 
         Height          =   285
         Left            =   5250
         TabIndex        =   10
         Top             =   1410
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox marka 
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   1395
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox weight 
         Height          =   315
         Left            =   7605
         TabIndex        =   17
         Top             =   2010
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox freight 
         Height          =   315
         Left            =   6075
         TabIndex        =   16
         Top             =   2010
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox bdated 
         Height          =   315
         Left            =   4920
         TabIndex        =   15
         Top             =   2010
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox biltno 
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         Top             =   2010
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox station 
         Height          =   315
         Left            =   150
         TabIndex        =   12
         Top             =   2010
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   780
         Left            =   180
         Top             =   6840
         Width           =   9600
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Total Discount : "
         Height          =   255
         Left            =   8070
         TabIndex        =   66
         Top             =   4800
         Width           =   1290
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Weight : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7605
         TabIndex        =   65
         Top             =   1755
         Width           =   1905
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Freight : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6120
         TabIndex        =   64
         Top             =   1755
         Width           =   1545
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bilty No. : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3360
         TabIndex        =   63
         Top             =   1755
         Width           =   1545
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Railway/Station : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   62
         Top             =   1755
         Width           =   1230
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Through : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2490
         TabIndex        =   61
         Top             =   1170
         Width           =   5775
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bundle(s) : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1140
         TabIndex        =   60
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3900
         TabIndex        =   59
         Top             =   465
         Width           =   1020
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Order By : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2115
         TabIndex        =   58
         Top             =   465
         Width           =   1635
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Gross Amount : "
         Height          =   255
         Left            =   6750
         TabIndex        =   57
         Top             =   4860
         Width           =   1260
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Net Amount : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6930
         TabIndex        =   56
         Top             =   6450
         Width           =   1155
      End
      Begin VB.Label label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Code : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4950
         TabIndex        =   55
         Top             =   465
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No. : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   54
         Top             =   465
         Width           =   1110
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   53
         Top             =   465
         Width           =   1020
      End
      Begin VB.Label mga 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6945
         TabIndex        =   52
         Top             =   6150
         Width           =   1125
      End
      Begin VB.Label mna 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8130
         TabIndex        =   51
         Top             =   6450
         Width           =   1200
      End
      Begin VB.Label mgd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8130
         TabIndex        =   50
         Top             =   6150
         Width           =   1200
      End
      Begin VB.Label tqu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3285
         TabIndex        =   49
         Top             =   6450
         Width           =   930
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Total Quantity : "
         Height          =   255
         Left            =   3390
         TabIndex        =   48
         Top             =   5010
         Width           =   1470
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Marka : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   47
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4845
         TabIndex        =   46
         Top             =   1755
         Width           =   1230
      End
      Begin VB.Label labelbybank 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   45
         Top             =   6450
         Width           =   885
      End
      Begin VB.Label labelbybanklbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "By Bank : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4275
         TabIndex        =   44
         Top             =   6450
         Width           =   795
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F4 Key To Delete A Invoive Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   43
         Top             =   6135
         Width           =   2895
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agent :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4980
         TabIndex        =   42
         Top             =   825
         Width           =   510
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transport"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1410
         TabIndex        =   41
         Top             =   1755
         Width           =   1935
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mark "
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8310
         TabIndex        =   40
         Top             =   1170
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmInvoice_blueprint"
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
Dim category As String
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
If Me.txtMark.Text = "M" Then
mystr1 = "MOHKAMPUR"
ElseIf Me.txtMark.Text = "W" Then
mystr1 = "W.K.ROAD"
ElseIf Me.txtMark.Text = "U" Then
mystr1 = "UTSAV COMPLEX"
End If

Dim LEFTM As Integer
Open "" + VB.App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
FooterYes = False
header:
    If kkk.State = 1 Then
          kkk.close
    End If
    CNSetup
    kkk.Open "select * from setup1", CON, adOpenStatic, adLockReadOnly, adCmdText
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
If rs1.State = 1 Then rs1.close
rs1.Open "select top 1000 * from invoicea_blue where " & stringyear & "", CON, adOpenStatic, adLockReadOnly
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); "To,   S.L. Code : "; Tab(20); Mid$(rs1!SUBLEDGER, 1, 5); Tab(50); "Invoice No. : "; Chr(27) + Chr(72); Trim(rs1!INVOICENO); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); rs1!invoicedate
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & kkk!DESCFORINVOICE
        Print #1, Tab(5); IIf(IsNull(kkk!address1), " ", kkk!address1); Tab(49); Chr(27) + Chr(71); "Order by    : "; Chr(27) + Chr(72); Trim(rs1!orderby); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!address2), " ", kkk!address2)
        Print #1, Tab(5); IIf(IsNull(kkk!address3), " ", kkk!address3); Tab(49); Chr(27) + Chr(71); "Bilty No.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        'Print #1, ""
        Print #1, Tab(73); Chr(27) + Chr(71); "(" & txtMark & ")"; Chr(27) + Chr(72)
        kkk.close
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
    If kk.State = 1 Then kk.close
    kk.Open "select * from invoiceb_blue where " & stringyear & " and invoiceno=" + Trim(rs1!INVOICENO) + " order by printorder,sno ", CON, adOpenStatic, adLockReadOnly, adCmdText
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
                vdis = kk!DISCOUNT
                tdata.Open "Select bookname from books where " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
                Print #1, Tab(0); rsets(Trim(Str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(Str(kk!quantity)), 5); Tab(58); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!quantity
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from invoiceb_blue where " & stringyear & " and invoiceno=" + Trim(rs1!INVOICENO) + " and printorder =" + Trim(Str(cdiscount)) + " group by printorder", CON, adOpenStatic, adLockReadOnly, adCmdText
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
       Print #1, Tab(50); rsets(Trim(Str(totalquantity)), 7); Tab(84); rsets(Trim(Format(Str(netamount), "0.00")), 12)
       Line = Line + 2
       If kk.State = 1 Then
             kk.close
       End If
       kk.Open "Select * from invoicec_blue where  " & stringyear & " and  invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
       If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(60); Trim(kk!Text) + " :  @  " + Trim(Format(Str(kk!rate), "0.00")) & " % "; Tab(84); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(60); Trim(kk!Text) & " :"; Tab(84); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
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
        kk.close
        Dim Va As Variant
        kk.Open "Select * from invoicea_blue where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
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
        tempdata.Open "setup1", CON, adOpenStatic, adLockReadOnly, adCmdTable
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
        PrintOption.Show
        
        

End Sub
Sub invoicecalc()
'OTHERSALES.calc
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     mna.Caption = Format(Round((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
End Sub
Sub invoiceabandon()
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
If Edit = False Then
  
  
  
    End If
        On Error Resume Next
        Dim ctl As Control
        For Each ctl In Me.Controls
            If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
                If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) Then
                    ctl.Text = ""
                End If
                ctl.Enabled = False
            End If
        Next
        For I = 1 To maxrow
           Grid1.Row = I
            For J = 1 To 8
                Grid1.Col = J
               Grid1.Text = ""
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
                Grid1.Text = tempmeb.Text
                '/*************************
                'If RS.State = 1 Then
                '    RS.close
                'End If
                'RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                Set RS = CON.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(Grid1.Text) & "'")
                Row = Grid1.Row
                Col = Grid1.Col
                If Trim(Grid1.Text) <> "" Then
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
                            Grid1.Text = RS(0)
                            Grid1.Col = 2
                            Grid1.Text = RS(1)
                         '   If Not edit Then
                                Grid1.Col = 3
                                If Trim(Grid1.Text) = "" Then
                                    Grid1.Text = 0
                                End If
                                q = Val(Grid1.Text)
                                Grid1.Col = 5
                                If Trim(Grid1.Text) = "" Then
                                    Grid1.Text = Format(RS(3), "0.00")            'rs(3)
                                    r = RS(3)
                              
                                End If
                                '/******************
                                category = returnCategory(Trim(RS(2)))
                                If category = "C1" Then
                                  Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
                                ElseIf category = "C2" Then
                                  Set kk = CON.Execute("select CATEGORY2 from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
                                ElseIf category = "C3" Then
                                  Set kk = CON.Execute("select CATEGORY3 from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
                                End If
                                
                                Grid1.Col = 6
                                If Grid1.Text = "" And addmode = True Then
                                    If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        kk.close
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "' and " & stringyear)
                                        Grid1.Col = 4
                                        If kk.BOF Then
                                             GoTo abc
                                        End If
                                        Grid1.Text = Format(kk(0), "0.00")
                                        Grid1.Col = 6
                                        Grid1.Text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = RS(3)
                                    Else
abc:
                                        Grid1.Col = 4
                                        Grid1.Text = Format(RS(4), "0.00")
                                        Grid1.Col = 6
                                        Grid1.Text = Format(RS(4), "0.00")
                                        D = RS(4)
                                End If
                            
                                Grid1.Col = 7
                                Grid1.Text = Format(Round(q * r, 2), "0.00")
                                Grid1.Col = 8
                                Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
                              Else
                              
                                  If Grid1.Text = "" And addmode = False Then
                                    If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        kk.close
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "' and " & stringyear)
                                        Grid1.Col = 4
                                        If kk.BOF Then
                                             GoTo abc
                                        End If
                                        Grid1.Text = Format(kk(0), "0.00")
                                        Grid1.Col = 6
                                        Grid1.Text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = RS(3)
                                    End If
                                  End If
                              
                              End If
                          '  End If
                            Grid1.Col = Col
                            RS.close
                        End If
                    End If
                End If
            Case 3, 5, 6
                If Grid1.Col <> 3 Then
                    Grid1.Text = Format(Trim(tempmeb.Text), "0.00")
                Else
                    Grid1.Text = Format(Trim(tempmeb.Text), "0")
                End If
                If Trim(Grid1.Text) = "" Then
                    Grid1.Text = 0
                End If
                Row = Grid1.Row
                Col = Grid1.Col
                Grid1.Col = 3
                q = Val(Trim(Grid1.Text))
                Grid1.Col = 5
                r = Val(Trim(Grid1.Text))
                Grid1.Col = 6
                D = Val(Trim(Grid1.Text))
                Grid1.Col = 7
                Grid1.Text = Format(Round(q * r, 2), "0.00")
                Grid1.Col = 8
                Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
                Grid1.Col = Col
            Case 4
                Grid1.Text = tempmeb.Text
                If Trim(Grid1.Text) = "" Then
                    Grid1.Text = 0
                End If
        End Select
        Row = Grid1.Row
        Col = Grid1.Col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.Col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.Text)), 2)
        Next
        invoicecalc
        Me.tqu.Caption = ""
        For I = 1 To maxrow
            Grid1.Col = 3
            Grid1.Row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
        Next
        Grid1.Row = RRR
        Grid1.Col = CCC
        templost = True
End Function

Private Sub bdated_LostFocus()
If Trim(bdated.Text) <> Trim("__/__/____") Then
   If Not checkdate(Trim(bdated.Text), bdated) Then
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
    If s1.State = 1 Then s1.close
    
    s1.Open "select category from [groups] where groupcode='" & s & "' and " & stringyear & "", CON
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
                If Trim(Bookname.Text) = "" Then
                    Grid1.Col = 1
                    If Trim(Grid1.Text) = "" Then
                        Grid1.Text = Bookname.Text
                           Bookname.SetFocus
  '********* vk
                          
                          
                          If Trim(Grid1.Text) = "" And Row = 1 Then
                                 Grid1.Col = 2
                                 Grid1.Text = ""
                                 If Trim(Grid1.Text) = "" Then
                                           
                                          Grid1.Col = 1
                                          Bookname.SetFocus
                                          Grid1.SetFocus
                                       Exit Sub
                                 End If
                           End If
              '********
                         If Commandother.Enabled = True Then
                           Commandother.SetFocus
                         End If
                           Exit Sub
                    End If
                End If
                Grid1.Row = Row
                Grid1.Col = Col
                Grid1.Text = Bookname.Text
                '/*************************
                If RS.State = 1 Then
                    RS.close
                End If
                RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                Row = Grid1.Row
                Col = Grid1.Col
                If Trim(Grid1.Text) <> "" Then
                    If Not RS.BOF Then
                        RS.MoveFirst
                        RS.Find "bookname='" + Trim(Grid1.Text) + "'"
                        If RS.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            RS.close
                            Exit Sub
                        Else
                            
                            Grid1.Col = 1
                            Grid1.Text = RS(0)
                            Grid1.Col = 2
                            Grid1.Text = RS(1)
                        '   If Not edit Then
                                 Grid1.Col = 3
                            If Trim(Grid1.Text) = "" Then
                                Grid1.Text = 0
                            End If
                            q = Val(Grid1.Text)
                            Grid1.Col = 5
                            Grid1.Text = Format(RS(3), "0.00")
                            r = RS(3)
                            '/******************
                            
                            category = returnCategory(Trim(RS(2)))
                            If category = "C1" Then
                            Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
                            ElseIf category = "C2" Then
                            Set kk = CON.Execute("select Category2 from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
                            ElseIf category = "C3" Then
                            Set kk = CON.Execute("select Category3 from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
                            
                            End If
                            
                            Grid1.Col = 6
                            
                            If Trim(kk(0)) <> "" Then
                               tempstr = Trim(kk(0))
                               kk.close
                               
                               Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "' and " & stringyear)
                               
                               Grid1.Col = 4
                               If kk.BOF Then
                                   GoTo abc
                               End If
                               Grid1.Text = Format(kk(0), "0.00")
                               Grid1.Col = 6
                               Grid1.Text = Format(kk(0), "0.00")
                               D = kk(0)
                            Else
abc:
                                 Grid1.Col = 4
                                 Grid1.Text = Format(RS(4), "0.00")
                                    Grid1.Col = 6
                                    Grid1.Text = Format(RS(4), "0.00")
                                    D = RS(4)
                                End If
                                Grid1.Col = 7
                                Grid1.Text = Round(q * r, 2)
                                Grid1.Col = 8
                                Grid1.Text = Round((q * r) * (D / 100), 2)
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
            totalamount = totalamount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.Col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.Text)), 2)
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

Private Sub Command1_Click()

End Sub
Private Sub cmbAgentName_GotFocus()
'cmbAgentName.ListIndex = 0
End Sub

Private Sub cmbAgentName_LostFocus()

If cmbAgentName.Text = "" Then
   MsgBox "Enter a Agent Name.. "
   cmbAgentName.SetFocus
   Exit Sub
Else
  Dim rs1 As New ADODB.Recordset
  rs1.Open "select *  from AgentMaster where AgentName='" & cmbAgentName.Text & "' and " & stringyear & " order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
  'If rs1.RecordCount <= 0 Then
  If rs1.EOF = True Then
     MsgBox "Enter valid Agent Name.. "
     cmbAgentName.SetFocus
  End If
End If

End Sub

Private Sub Commandprintnh_Click()
    printch = "invoicea_blue"
    ino = I_NO
    printch1 = "INVOICENO"


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
    Set RS = New ADODB.Recordset
    Dim TEMPNUM As Integer
    
    If Edit = False Then
       Me.I_NO.Text = CON.Execute("Select max(invoiceno) from invoicea_blue where " & stringyear)(0) + 1
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
    
    Me.Edit = False
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
    Grid1.Enabled = True
    Me.customercode.Enabled = True
    
    addoredit = True
    
    I_NO.SetFocus
    
End Sub
Private Sub Commandall_Click()
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Dim myvalue As String

If Trim(Me.customercode.Text) = "" Then
    MsgBox "Please Fill the customer detail "
    Exit Sub
End If

myvalue = InputBox("Please enter the quantity ", "Enter the quantity: ", "1")
    
If Len(myvalue) > 0 And Val(myvalue) > 0 Then
    
    
    
    Grid1.Rows = 1
    Grid1.Rows = 2
    Grid1.Col = 1
    Grid1.Row = 1
    If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select * from books order by BOOKCODE", CON, adOpenDynamic, adLockReadOnly, adCmdText
    Row = Grid1.Row
    Col = Grid1.Col
    If Not RS.BOF Then
        RS.MoveFirst
        Do While Not RS.EOF
            Grid1.Col = 1
            Grid1.Text = RS(0)
            Grid1.Col = 2
            Grid1.Text = RS(1)
            Grid1.Col = 3
            If Trim(Grid1.Text) = "" Then
                Grid1.Text = Val(myvalue)
            End If
            q = Val(Grid1.Text)
            Grid1.Col = 5
            Grid1.Text = Format(RS(3), "0.00")
            r = RS(3)
            
            '/******************
            'Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
            
                          
            category = returnCategory(Trim(RS(2)))
            If category = "C1" Then
               Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
            ElseIf category = "C2" Then
               Set kk = CON.Execute("select Category2 from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
            ElseIf category = "C3" Then
               Set kk = CON.Execute("select Category3 from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
            
            End If
             
            Grid1.Col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.close
                Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "' and " & stringyear)
                Grid1.Col = 4
                If kk.BOF Then
                    GoTo abc
                End If
                Grid1.Text = Format(kk(0), "0.00")
                Grid1.Col = 6
                Grid1.Text = Format(kk(0), "0.00")
                D = kk(0)
            Else
abc:
                Grid1.Col = 4
                Grid1.Text = Format(RS(4), "0.00")
                Grid1.Col = 6
                Grid1.Text = Format(RS(4), "0.00")
                D = RS(4)
            End If
            Grid1.Col = 7
            Grid1.Text = Format(Round(q * r, 2), "0.00")
            Grid1.Col = 8
            Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
            If Not RS.EOF Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Row = Grid1.Row + 1
                RS.MoveNext
            End If
        Loop
    
    '/**fghfghgh
    
    End If
    RS.close
    totalamount = 0
    totaldiscount = 0
    Me.tqu.Caption = ""
    For I = 1 To Grid1.Rows - 1
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.Col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.Col = 3
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
     Next
     maxrow = Grid1.Rows - 1
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
   
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select top 1 * from invoicea_blue where invoiceno=" & I_NO.Text & " and " & stringyear, CON, adOpenStatic, adLockReadOnly
    If rs1.EOF = False Then
        If rs1!bAuthorized = True Then
            MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
            Exit Sub
        End If
       
    End If


'=======================================




If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
                CON.Execute ("delete  from invoicea_blue where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete  from invoiceb_blue where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete  from invoicec_blue where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                invoiceabandon
End If




Exit Sub
Del:
MsgBox "" & err.DESCRIPTION

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
    Grid1.Enabled = True
    Commandall.Enabled = False
    Me.customercode.Enabled = True
    Edit = True
    addoredit = False
    I_NO_LostFocus
    i_dt.Enabled = True
    i_dt.SetFocus
    
    CON.Execute ("delete  from invoicectmp_blue WHERE " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
    DoEvents
    CON.Execute ("insert into invoicectmp_blue([INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],[saleType]) " & _
    "  select [INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],[saleType]" & _
    " from invoicec_blue where  " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
    DoEvents
    
    Dim kx As Integer
    kx = 0
    '''Do While kx < 18000
    '''kx = kx + 1
    '''Loop
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
    
    
    
    
End Sub
Private Sub Commandother_Click()
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
Commandsave.Enabled = True

searchForm = "invoiceblue"
frmEndPartTrans.Show
frmEndPartTrans.Refresh
DoEvents
DoEvents
DoEvents
DoEvents

 
End Sub
Private Sub CommandPrint_Click()
  
  printch = "invoicea_blue"
  ino = I_NO
  printch1 = "INVOICENO"
  
  Printheader = True
  printinvoice
   
End Sub
Private Sub Commandreturn_Click()
   
'''   Dim RS As New ADODB.Recordset
'''   RS.Open "tempINV", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
'''   If RS.BOF Then
'''       RS.AddNew
'''   End If
'''   RS!In = CON.Execute("Select max(invoiceno) from INVOICEA")(0)
'''   RS.Update
'''   RS.Close
'''
   Unload Me
'''   addoredit = False
'''   'MainMenu.Toolbar1.Visible = True
End Sub

Private Sub Commandsave_Click()
    
On Error GoTo save_
    
    
    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
  
   
  
    If rs1.State = 1 Then rs1.close
    rs1.Open "select top 100 * from invoicea_blue where invoiceno=" & I_NO.Text & " and " & stringyear, CON, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select top 100 * from invoicea_blue where invoiceno=" & I_NO.Text & " and " & stringyear, CON, adOpenKeyset, adLockReadOnly
       'If rs_h.Fields("Print_yes").Value = "y" Then
          If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
          End If
       'End If
       
    End If
    
        
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
    If Trim(Grid1.Text) = "" Then
       MsgBox "Please Enter item.... "
       Exit Sub
    End If
    SAVED = False
    
    
    If Edit = False Then
       If check_Duplikate("invoicea_blue", I_NO.Text) = True Then
           If CON.Execute("Select max(invoiceno) from invoicea_blue")(0) >= Val(Trim(Me.I_NO.Text)) Then
                Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
           End If
         'Exit Sub
       End If
    End If
    
    
    If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(customercode.Text) <> "" Then
            
            If Edit Then
                
                CON.Execute ("delete  from invoicea_blue where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete  from invoiceb_blue where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete  from invoicec_blue where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                
            End If
            
            If RS.State = 1 Then
                RS.close
            End If
            LAMOUNT = 0
            
    'Code for Order Mnm
    
    If Edit = False Then
          'If (I_OB <> "" And txtMark <> "") Then
          'Party_Remove_FromOrder Trim(Me.customercode.Text), txtMark, Trim(I_OB)
          'End If
    End If
    
            
    RS.Open "select * from invoicea_blue where " & stringyear & " and invoiceno <=0", CON, adOpenDynamic, adLockOptimistic
    If Not Edit Then
again:
    End If
            
            
            
            RS.AddNew
            RS!INVOICENO = Val(Me.I_NO.Text)
            RS!invoicedate = Me.i_dt.Text
            RS!Genledger = Trim(Me.Genledger.Text)
            RS!SUBLEDGER = Trim(Me.customercode.Text)
            RS!agentname = Trim(Me.cmbAgentName.Text)
            RS!transportname = Trim(Me.cmbtransportname.Text)
            RS!orderby = Trim(Me.I_OB.Text)
            If Trim(Me.I_DTOB) <> Trim("__/__/____") Then
            '    rs!ORDERDATE = Date
            'Else
                RS!ORDERDATE = Trim(Me.I_DTOB.Text)
            End If
            RS!marka = Trim(Me.marka.Text)
            RS!Godown = IIf(txtMark.Text = "", "n", txtMark.Text)
            RS!bundles = Trim(Me.bundles)
            RS!through = Trim(Me.through.Text)
            RS!through1 = Trim(Me.through1.Text)
            If Trim(Me.through1.Text) = "" Then
                RS!through1 = " "
            End If
            RS!station = Trim(Me.station.Text)
            RS!biltyno = Trim(Me.biltno.Text)
            If Trim(Me.bdated) <> Trim("__/__/____") Then
                RS!BILTYDATE = Me.bdated & ""
           End If
            
            RS!freight = Trim(Me.freight)
            RS!weight = Trim(Me.weight)
            RS!netamount = Round(Val(Trim(Me.mna.Caption)), 2)
            RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
            RS!txt1 = Trim(frmEndPartTrans.T1TEXT.Text)
            RS!txt1a = Val(Trim(frmEndPartTrans.T1.Text))
            RS!txt2 = Trim(frmEndPartTrans.T2TEXT.Text)
            RS!txt2a = Val(Trim(frmEndPartTrans.T2.Text))
            RS!baa = Val(Trim(frmEndPartTrans.T3TEXT.Text))
            RS!baa = Val(Trim(labelbybank.Caption))
            If addmode = True Then
                If Val(Trim(frmEndPartTrans.T3TEXT.Text)) <> 0 Then
                      RS!advicestatus = "Pending"
                      Me.txtadst.Text = "Pending"
                End If
            Else
                RS!advicestatus = Me.txtadst.Text & ""
            End If
            Dim trs As New ADODB.Recordset
            trs.Open " SELECT DISTCODE FROM SLEDGER  WHERE SUBLEDGER='" & customercode.Text & "' and " & stringyear, CON, adOpenStatic, adLockOptimistic, adCmdText
            If Not trs.BOF Then
                RS!District = Trim(trs!distcode)
            Else
                RS!District = ""
            End If
err1:
           If Not Edit Then
                If CON.Execute("Select max(invoiceno) from invoicea_blue where " & stringyear & "")(0) >= Val(Trim(Me.I_NO.Text)) Then
                    On Error GoTo err1
                End If
            End If
            RS!fyear = session
            RS!setupid = setupid
            RS.update
            
            
            On Error GoTo 0
            RS.close
            RS.Open "select * from invoiceb_blue where " & stringyear & " and invoiceno<=0", CON, adOpenDynamic, adLockOptimistic
            Dim I As Integer
            RRRR = Grid1.Row
            CCCC = Grid1.Col
            For I = 1 To maxrow
                Grid1.Row = I
                Grid1.Col = 1
                If Trim(Grid1.Text) <> "" Then
                    Grid1.Col = 3
                    If Val(Trim(Grid1.Text)) > 0 Then
                       Grid1.Col = 5
                       If Val(Trim(Grid1.Text)) > 0 Then
                         RS.AddNew
                         Grid1.Col = 1
                         RS!INVOICENO = Val(Me.I_NO.Text)
                         RS!invoicedate = Me.i_dt.Text
                         RS!Genledger = Trim(Me.Genledger.Text)
                         RS!SUBLEDGER = Trim(Me.customercode.Text)
                         RS!Bookcode = Trim(Grid1.Text)
                         Grid1.Col = 3
                         RS!quantity = Trim(Grid1.Text)
                         Grid1.Col = 5
                         RS!rate = Trim(Grid1.Text)
                         Grid1.Col = 7
                         RS!amount = Trim(Grid1.Text)
                         LAMOUNT = Val(Trim(Grid1.Text))
                         Grid1.Col = 4
                         RS!PRINTORDER = Trim(Grid1.Text)
                         Grid1.Col = 6
                         RS!DISCOUNT = Trim(Grid1.Text)
                         Grid1.Col = 8
                         RS!netamount = LAMOUNT - Trim(Grid1.Text)
                         LAMOUNT = 0
                         RS!agentname = Trim(Me.cmbAgentName.Text)
                         
                        RS!fyear = session
                        RS!setupid = setupid
                         
                         RS.update
                       End If
                    End If
                End If
            Next
            RS.close
            Grid1.TopRow = 1
            Grid1.Row = 1
            Grid1.Col = 1
            RS.Open "select * from invoicec_blue where " & stringyear & " and invoiceno<=0", CON, adOpenDynamic, adLockOptimistic
            '/******
                'Dim I, x As Integer
                Dim temprs As ADODB.Recordset
                Set temprs = New ADODB.Recordset
                   
                With frmEndPartTrans
                
                
                For I = 1 To .vs.Rows - 1
                   
                   If Trim(.vs.TextMatrix(I, 0)) <> "" Then
                   
                        
                        RS.AddNew
                        RS!fyear = session
                        RS!setupid = setupid
 
                        RS!INVOICENO = Val(Me.I_NO.Text)
                        RS!invoicedate = Me.i_dt.Text
                        RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
                        RS!Text = Trim(.vs.TextMatrix(I, 0))
                        If temprs.State = 1 Then
                            temprs.close
                        End If
                        
                        
                        If Edit Then
                        temprs.Open "select * from invoicectmp_blue WHERE  INVOICENO=" & frmInvoice_blueprint.I_NO & " and " & stringyear & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
                        If .vs.TextMatrix(I, 0) <> "" Then
                                temprs.Find "TEXT='" + Trim(.vs.TextMatrix(I, 0)) + "'"
                                RS!Genledger = Trim(temprs!Genledger)
                                RS!SUBLEDGER = Trim(temprs!SUBLEDGER)
                                RS!DebitorCredit = Trim(temprs!DebitorCredit)
                                RS!RYN = temprs!RYN & ""
                                
                        End If
                        temprs.close
                        
                        
                        
                Else
                        
                        temprs.Open "select * from INVOICEEND where  type='invoiceblue' and " & stringyear & " order by printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
                        If .vs.TextMatrix(I, 0) <> "" Then
                                temprs.Find "TEXT='" + Trim(.vs.TextMatrix(I, 0)) + "'"
                                RS!Genledger = Trim(temprs!Genledger)
                                RS!SUBLEDGER = Trim(temprs!SUBLEDGER)
                                RS!DebitorCredit = Trim(temprs!DebitorCredit)
                                RS!RYN = temprs!RYN & ""
                        End If
                        temprs.close
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
                 
            RS.close
                

            SAVED = True
        
        End If
             
            s11 = ""
            ss11 = ""
            
            s11 = InStr(1, Me.station.Text, " ")
            If s11 <> 0 Then
            ss11 = Trim(Mid(Me.station.Text, 1, s11))
            Else
            ss11 = Me.station.Text
            End If
            PopUpValue1 = ss11
   
             
             UpdateDisPatchReg I_NO, i_dt, Me.customercode, PopUpValue1, Trim(Me.bundles), Trim(Me.cmbtransportname.Text), Trim(Me.marka.Text), Trim(Me.biltno.Text), Me.bdated, Trim(Me.freight), "DispatchRegister"
             PopUpValue1 = ""
         ' End If

        
        If SAVED Then
            MsgBox "Record Saved"
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
            Me.Commandprintnh.Enabled = True
        End If
        addmode = False
        addoredit = False
        SetButton Commandadd, Commandedit, Commandsave, Commanddelete
        Me.Commandsave.Enabled = False


Exit Sub
save_:
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub Commandsearch_Click()


sqlqry = "select InvoiceNo,InvoiceDate,Subledger,NetAmount from invoicea_blue where " & stringyear & "  InvoiceNo"
orderby = "order by InvoiceNo"


searchType = "inv"
popuplist10 "select InvoiceNo,InvoiceDate,Subledger,NetAmount from invoicea_blue where " & stringyear & "  order by InvoiceNo", CON



End Sub
Private Sub Commandsearch_GotFocus()
  
If PopUpValue1 <> "" Then

  'If Val(inviceNo) > 0 Then
     I_NO.Text = PopUpValue1
     
     PopUpValue1 = ""
     I_NO_LostFocus
     
     
  'End If
  

End If
  
  
End Sub
Private Sub customercode_LostFocus()
    
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "select * from sledger where gledger='SUNDRY DEBTORS' and subledger='" + Trim(customercode.Text) + "' and " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
    'If RS.RecordCount <= 0 Then
    If (RS.EOF = True Or RS.RecordCount <= 0) Then
        customercode.SetFocus
        HIT
        RS.close
        Exit Sub
    End If
    
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    If RS!distcode <> "" And addmode = True Then
       rs1.Open "Select * from Districts where Districtname = '" & RS!distcode & "' and " & stringyear, CON, adOpenStatic, adLockReadOnly
       If rs1.RecordCount > 0 Then
          Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
       End If
    End If
    RS.close
    Me.textbox.Text = Me.customercode.Text
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
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If Grid1.Row >= 1 Then
           Grid1.RemoveItem Grid1.Row
           a = Grid1.Text
           tempmeb.Text = a
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
                 SendKeys "{tab}"
                 Exit Sub
            End If
             SendKeys "{DOWN}"
             SendKeys "{TAB}"
        Else
            If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("grid1")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("tempmeb")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bookname")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("weight")) Then
                SendKeys ("{TAB}")
            End If
        End If
    End If
    
    
End Sub
Private Sub Form_Load()
    
Me.Top = -10
Me.Left = 50

Me.Width = 10000
Me.Height = 7800
    
    
    
Me.Caption = "Invoice (Bluprint)"
    
    
Screen.MousePointer = vbHourglass
    
    
Dim rs_godwn As New ADODB.Recordset

If rs_godwn.State = 1 Then rs_godwn.close
rs_godwn.Open "select * from GodownMaster where len(Godwn)<=3 and " & stringyear & " order by id", CON, adOpenForwardOnly, adLockReadOnly
txtMark.Clear
If Not rs_godwn.EOF Then
Do While Not rs_godwn.EOF
   If IsNull(rs_godwn(0)) = False Then
     Me.txtMark.AddItem rs_godwn(0)
   End If
   If Not rs_godwn.EOF Then rs_godwn.MoveNext
 Loop
End If



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
Me.Top = 50
Me.Left = 50
Grid1.Rows = 2
Grid1.Cols = 1
Grid1.Rows = 10
Grid1.Cols = 9
Grid1.Row = 0
Grid1.Col = 1
Grid1.Text = "Book Code "
Grid1.Col = Grid1.Col + 1
Grid1.Text = "Book Name"
Grid1.Col = Grid1.Col + 1
Grid1.Text = "Quantity"
Grid1.Col = Grid1.Col + 1
Grid1.Text = "Print. Ord."
Grid1.Col = Grid1.Col + 1
Grid1.Text = "Rate"
Grid1.Col = Grid1.Col + 1
Grid1.Text = "Disc %"
Grid1.Col = Grid1.Col + 1
Grid1.Text = "Amount"
Grid1.Col = Grid1.Col + 1
Grid1.Text = "Disc. Amount"
Grid1.RowHeight(0) = Grid1.CellHeight + 50
Grid1.ColWidth(0) = 150
Grid1.ColWidth(1) = 1000
Grid1.ColWidth(2) = 2500
Grid1.ColWidth(3) = 750
Grid1.ColWidth(4) = 750
Grid1.ColWidth(5) = 850
Grid1.ColWidth(6) = 800
Grid1.ColWidth(7) = 1150
Grid1.ColWidth(8) = 1200
Bookname.Height = 2325
Me.CommandPrint.Enabled = True
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
RS.close
    
    
    Genledger.Text = "SUNDRY DEBTORS"
    'Set RS = CON.Execute("exec fatch_ledger '" & Genledger.Text & "','" & session & "'," & main.setupid & "")
    RS.Open "select * from sledger where gledger='" & Genledger.Text & "' and  " & stringyear, CCON, adOpenDynamic, adLockReadOnly, adCmdText
    
    If Not RS.BOF Then
        Do While Not RS.EOF
            Me.customercode.AddItem RS("subledger")
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    
     '*******Agent  combo fill
    RS.Open "select  Agentname from AgentMaster where " & stringyear & " order by agentname", CON_blue, adOpenDynamic, adLockReadOnly, adCmdText
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
    RS.Open "select  transportname from transportMaster order by transportname", CON, adOpenDynamic, adLockReadOnly, adCmdText
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




    On Error Resume Next

    Bookcode.Left = Grid1.Left
    Bookcode.Visible = False
    Bookname.Visible = False
    Grid1.Rows = 100
    For I = 1 To 99
        Grid1.RowHeight(I) = 300
    Next
    
    Bookcode.Width = 1230
    Bookname.Width = 2830
    amount.Width = rate.Width

       If kk.State = 1 Then kk.close
       kk.Open "SELECT MAX(INVOICENO) FROM invoicea_blue where " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
       If kk(0) <> "" Then
           addoredit = False
          Me.I_NO.Text = kk(0)
          I_NO_LostFocus
       Else
          Me.I_NO.Text = "1"
       End If

    
    

    Commanddelete.Enabled = True
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
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete

  
 Screen.MousePointer = vbDefault
    
    
 BackColorFrom Me, 1
 
'

End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub

Private Sub Grid1_Click()
If Trim(Me.customercode.Text) <> "" Then
Dim PREVROW As Integer
Dim prevcol As Integer
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
prevcol = Grid1.Col
PREVROW = Grid1.Row
If Grid1.Row > 1 Then
    Grid1.Row = Grid1.Row - 1
    Grid1.Col = 1
    If Trim(Grid1.Text) <> "" Then
        Grid1.Row = PREVROW
        Grid1.Col = prevcol
        If Trim(Me.customercode.Text) <> "" Then
            If Me.customercode.Enabled = True Then
                Me.customercode.Enabled = False
            End If
            Grid1.Col = 1
            If prevcol > 1 And Trim(Grid1.Text) = "" Then
                Grid1.Col = 2
                SendKeys Chr(13)
            Else
                Grid1.Col = prevcol
                SendKeys Chr(13)
            End If
        Else
            MsgBox "Please fill the customer detail first"
        End If
    End If
Else
    If Trim(Me.customercode.Text) <> "" Then
        If Me.customercode.Enabled = True Then
            Me.customercode.Enabled = False
        End If
        Grid1.Col = 1
        If prevcol > 1 And Trim(Grid1.Text) = "" Then
            Grid1.Col = 2
            Grid1.SetFocus
            SendKeys Chr(13)
        Else
        'IF GRID1.COL
            Grid1.Col = prevcol
            Grid1.SetFocus
            SendKeys Chr(13)
        End If
        'SendKeys Chr(13)
    End If
End If
End If
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
If Trim(Me.customercode.Text) <> "" Then
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
                        tempmeb.Text = Format(Grid1.Text, "0.00")
                        
                    Else
                        tempmeb.Text = Format(Grid1.Text, "0")
                    End If
                   
                Else
                    tempmeb.Text = Grid1.Text
                End If
                tempmeb.Width = Grid1.ColWidth(Grid1.Col)
                tempmeb.Left = Grid1.CellLeft + leftAlign
                tempmeb.Top = Grid1.Top + Grid1.CellTop '- 50
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.Text = Grid1.Text
                Bookname.Top = Grid1.Top + Grid1.CellTop
                Bookname.Left = Grid1.CellLeft + leftAlign
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
                    SendKeys Chr(KeyAscii)
                End If
            Case 1, 3, 4, 5, 6
                mprevcol = Grid1.Col
                tempmeb.SetFocus
            Case Else
                If KeyAscii = 13 Then
                    SendKeys "{RIGHT}"
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
   PopupMenu dd, , Grid1.Left + X, Grid1.Top + Y
End If
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

If Not IsDate(i_dt.Text) Then
i_dt.SetFocus
Exit Sub
End If


End If

End Sub

Private Sub I_DTOB_LostFocus()
If Trim(I_DTOB.Text) <> "__/__/____" Then
    If Not checkdate(Trim(I_DTOB.Text), I_DTOB) Then
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
Sub cmdButtonLock()
    Commandother.Enabled = False
    Commandall.Enabled = False
    Commandadd.Enabled = False
    Commandedit.Enabled = False
    Commandsearch.Enabled = False
    Commanddelete.Enabled = False
    Commandabandon.Enabled = False
    Commandprintnh.Enabled = True
    CommandPrint.Enabled = True
End Sub
Sub I_NO_LostFocus()

On Error Resume Next

Dim rs1 As ADODB.Recordset

If Val(inviceNo) > 0 Then
   I_NO.Text = inviceNo
   cmdButtonLock
End If



inviceNo = ""

Set rs1 = New ADODB.Recordset
Set RS = New ADODB.Recordset

    If Trim(I_NO.Text) = "" Then
        MsgBox "Invoice cannot be null"
        I_NO.SetFocus
    Else
        
        If RS.State = 1 Then
           RS.close
        End If
        RS.Open "Select top 1 * from  invoicea_blue where INVOICENO = " + Trim(I_NO.Text) + " and " & stringyear, CON, adOpenStatic, adLockReadOnly
        If RS.EOF = True Then
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
        
'        Dim ctl As Control
'        For Each ctl In Me.Controls
'            If Not TypeOf ctl Is CommandButton Then
'                ctl.Enabled = True
'            End If
'        Next

        I_NO.Text = RS!INVOICENO
        Me.i_dt.Text = RS!invoicedate
        Me.Genledger.Text = Trim(RS!Genledger)
        Me.customercode.Text = Trim(RS!SUBLEDGER)
        Me.cmbAgentName.Text = IIf(IsNull(RS!agentname), "", RS!agentname)
        Me.cmbtransportname.Text = IIf(IsNull(RS!transportname), "", RS!transportname)
        Me.textbox.Text = Trim(RS!SUBLEDGER)
        Me.I_OB.Text = IIf(IsNull(RS!orderby), "", Trim(RS!orderby))
        If RS!ORDERDATE <> "" Then
        Me.I_DTOB.Text = RS!ORDERDATE
        End If
        Me.marka.Text = IIf(IsNull(RS!marka), "", Trim(RS!marka))
        txtMark.Text = RS!Godown & ""
        Me.bundles = IIf(IsNull(RS!bundles), "", RS!bundles)
        Me.through.Text = IIf(IsNull(RS!through), "", RS!through)
        Me.through1.Text = IIf(IsNull(RS!through1), "", RS!through1)
        Me.station.Text = IIf(IsNull(RS!station), "", RS!station)
        Me.biltno.Text = IIf(IsNull(RS!biltyno), "", RS!biltyno)
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
        RS.close
       
       ' OTHERSALES.Form_Load
       '*/**/*/*/*/*//*/*
       
        If RS.State = 1 Then
           RS.close
        End If
        
       CON.Execute "select * from INVOICEctmp_blue WHERE INVOICENO=" & frmInvoice_blueprint.I_NO & " and " & stringyear
       RS.Open "Select * from invoiceb_blue where INVOICENO =" + Trim(I_NO.Text) + " and " & stringyear & " order by SNO", CON, adOpenStatic, adLockReadOnly
       Grid1.TopRow = 2
        If Not RS.EOF Then
        
            Grid1.Row = 1
            Grid1.Col = 1
            Do While Not RS.EOF
            aa = RS.RecordCount
               If Trim(RS!INVOICENO) = Trim(I_NO.Text) Then
                Grid1.Col = 1
                Grid1.Text = Trim(RS!Bookcode)
                If kk.State = 1 Then
                    kk.close
                End If
                kk.Open "select * from books where bookcode='" + Trim(RS!Bookcode) + "' and " & stringyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
                Grid1.Col = 2
                Grid1.Text = Trim(kk!Bookname)
                Grid1.Col = 3
                Grid1.Text = Trim(RS!quantity)
                Grid1.Col = 5
                Grid1.Text = Format(Round(RS!rate, 2), "0.00")
                Grid1.Col = 7
                Grid1.Text = Format(Round(RS!amount, 2), "0.00")
                Grid1.Col = 4
                Grid1.Text = Format(Round(RS!PRINTORDER, 2), "0.00")
                Grid1.Col = 6
                Grid1.Text = Format(Round(RS!DISCOUNT, 2), "0.00")
                Grid1.Col = 8
                Grid1.Text = Format(Round(RS!amount * (RS!DISCOUNT / 100), 2), "0.00")
                End If
                If Not RS.EOF Then
                    RS.MoveNext
                End If
                Grid1.Row = Grid1.Row + 1
                Grid1.Rows = Grid1.Rows + 1
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
            totalamount = totalamount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.Col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.Text)), 2)
        Next
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     Me.tqu.Caption = ""
        For I = 1 To maxrow
            Grid1.Col = 3
            Grid1.Row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
        Next
        Grid1.Row = RRR
        Grid1.Col = CCC
       ' templost = True
    End If
    Me.Commandother.Enabled = True
    
    
     If Val(inviceNo) > 0 Then
      I_NO.Text = inviceNo
      cmdButtonLock
   End If
   
   DoEvents
   DoEvents
   DoEvents
   DoEvents
   Grid1.Redraw = True
   
   Commandother.Enabled = False
   
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
If Grid1.Col = 1 Or Grid1.Col = 2 Then
    Grid1.Text = tempmeb.Text
Else
    If Grid1.Col = 3 Then
        Grid1.Text = Format(tempmeb.Text, "0")
    Else
        Grid1.Text = Format(tempmeb.Text, "0.00")
    End If
End If
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
                    Set RS = CON.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(Grid1.Text) & "'")
                    'RS.Open "books", CON, adOpenStatic, adLockReadOnly, adCmdTable
                    If Not RS.BOF Then
                        'RS.MoveFirst
                        'RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF And Trim(Grid1.Text) <> "" Then
                            RS.close
                            Exit Sub
                        Else
                            RS.close
                        If Trim(Grid1.Text) <> "" Then
                                Grid1.Col = 3
                            Else
                                Grid1.Col = 2
                            End If
                        End If
                    Else
                        If Trim(Grid1.Text) <> "" Then
                            Grid1.Col = 3
                        Else
                            Grid1.Col = 2
                        End If
                    End If
                    Grid1.SetFocus
                    Grid1_Click
                Case 3
                    If Val(tempmeb.Text) > 0 Then
                        Grid1.Col = Grid1.Col + 2
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                Case 4
                    Grid1.Col = Grid1.Col + 2
                    Grid1.SetFocus
                    Grid1_Click
                Case 5
                    If Val(tempmeb.Text) > 0 Then
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
                    Grid1.Rows = Grid1.Rows + 1
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
    

Exit Sub
aa11:
MsgBox "" & err.DESCRIPTION
    
End Sub
Private Sub tempmeb_LostFocus()
    If templost Then
        tempmeb.Visible = False
    End If
End Sub
Private Sub textbox_GotFocus()
    Me.customercode.Enabled = True
    Me.customercode.Visible = True
  '  Me.customercode.Height = 1100
    Me.customercode.ZOrder
    Me.customercode.SetFocus
    
End Sub
Private Sub through_LostFocus()
through = UCase(through)
End Sub
Private Sub through1_LostFocus()
through1 = UCase(through1)
End Sub

Private Sub weight_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(Me.customercode.Text) <> "" Then
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
            kkk.close
      End If
      If flagyes = True Then
      CNSetup
          kkk.Open "select * from setup1", CON, adOpenDynamic, adLockReadOnly, adCmdText
          If Not kkk.BOF Then
             Print #1, Chr(27) + Chr(18) + Chr(14)
             Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(18) + Chr(14); Trim(kkk!cname)
             Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(18); dspace(Trim(kkk!add1))
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
       Print #1, Chr(27) + Chr(18)
       Line = Line + 4
  End If
  
  
  
    If rs1.State = 1 Then
        rs1.close
    End If
    rs1.Open "invoicea_blue", CON, adOpenDynamic, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
   
    Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!SUBLEDGER; Tab(T5); "Invoice No. : "; Trim(rs1!INVOICENO); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!invoicedate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.close
            End If
            kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "' and " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE;
                Print #1, Tab(3); kkk!address1; Tab(T5); "Order by    : "; Trim(rs1!orderby); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!ORDERDATE
                Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.   : "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!BILTYDATE
                
                
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
            kk.Open "select * from invoiceb_blue where invoiceno=" + Trim(rs1!INVOICENO) + " and " + stringyear + " order by discount,printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kk.BOF Then
                kk.MoveFirst
                Dim cdiscount As Double
                Dim sno As Integer
                Dim tdata As ADODB.Recordset
                Set tdata = New ADODB.Recordset
                sno = 1
                Do While Not kk.EOF
                    cdiscount = kk!DISCOUNT
                    Do While kk!DISCOUNT = cdiscount
                        tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "' and " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
                        Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(Str(kk!quantity)), 5); Tab(T6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                        totalquantity = totalquantity + kk!quantity
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
                        tdata.Open "select sum(amount) from invoiceb_blue where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(Str(cdiscount)) + " and " & stringyear & " group by discount", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
           kk.Open "Select * from invoicec_blue where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
           If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5); Trim(kk!Text) + "    " + Trim(Format(Str(Round(kk!rate, 2)), "0.00")); Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5); Trim(kk!Text); Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
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
           kk.Open "Select * from invoicea_blue where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
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
            tempdata.Open "setup1", CON, adOpenDynamic, adLockReadOnly, adCmdTable
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
      kkk.close
End If
CNSetup
kkk.Open "select * from setup1", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
   rs1.close
End If
'Print #1, Chr(27) + Chr(14)
'line = line + 1
If rs1.State = 1 Then
    rs1.close
End If
rs1.Open "invoicea_blue", CON, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To, S.L. Code : "; Tab(T1 + 10); Mid$(rs1!SUBLEDGER, 1, 5); Tab(T5); "Invoice No. : "; Trim(rs1!INVOICENO); Tab(T8); "Dated     : "; rs1!invoicedate   'Chr(27) + Chr(18);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "' and " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE
        Print #1, Tab(3); kkk!address1; Tab(T5); "Order by    : "; Trim(rs1!orderby); Tab(T8); "Dated     : "; IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
        Print #1, Tab(3); kkk!address2; Tab(T5); "Bilty No.   : "; Trim(rs1!biltyno); Tab(T8); "Dated     : "; IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, Tab(3); kkk!address3
        kkk.close
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
        kk.close
    End If
    kk.Open "select * from invoiceb_blue where invoiceno=" + Trim(rs1!INVOICENO) + " and " & stringyear & " order by discount,printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
        kk.MoveFirst
        Dim cdiscount As Double
        Dim sno As Integer
        Dim tdata As ADODB.Recordset
        Set tdata = New ADODB.Recordset
        sno = 1
        Do While Not kk.EOF
            cdiscount = kk!DISCOUNT
            Do While kk!DISCOUNT = cdiscount
                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "' and " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
                Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(Str(kk!quantity)), 5); Tab(T6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!quantity
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
                tdata.Open "select sum(amount) from invoiceb_blue where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(Str(cdiscount)) + " and " & stringyear & " group by discount", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
        Print #1, repli("-", 145)
        Print #1, Tab(T5 - 4); rsets(Trim(Str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12)
       'by vk Print #1, Tab(t8); repli("-", 22)
        Line = Line + 2
        If kk.State = 1 Then
             kk.close
        End If
        kk.Open "Select * from invoicec_blue where " & stringyear & " and  invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5 + 21); Trim(kk!Text) + " :  @  " + Trim(Format(Str(Round(kk!rate, 2)), "0.00")) & " % "; Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5 + 20); Trim(kk!Text) & " :"; Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
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
        kk.close
        kk.Open "Select * from invoicea_blue where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
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
        tempdata.Open "setup1", CON, adOpenDynamic, adLockReadOnly, adCmdTable
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





