VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBookIssueSp_Ret 
   ClientHeight    =   10056
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10056
   ScaleWidth      =   13020
   Begin VB.Frame panel 
      Caption         =   "Book Return  (Specimen)"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   9936
      Left            =   135
      TabIndex        =   15
      Top             =   60
      Width           =   12840
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
         TabIndex        =   60
         Top             =   9432
         Width           =   984
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
         TabIndex        =   59
         Top             =   4320
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.ComboBox cboGodown 
         Height          =   288
         Left            =   9330
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1428
         Width           =   840
      End
      Begin VB.ComboBox cmbtransportname 
         Height          =   315
         Left            =   4380
         TabIndex        =   7
         Top             =   1395
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtadst 
         Height          =   315
         Left            =   6510
         TabIndex        =   31
         Top             =   7935
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ComboBox cmbAgentName 
         Height          =   288
         Left            =   2340
         TabIndex        =   2
         Top             =   684
         Width           =   4470
      End
      Begin VB.CommandButton Commandall 
         Caption         =   "All Books"
         Height          =   465
         Left            =   1395
         TabIndex        =   30
         Top             =   7710
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton Commandother 
         Caption         =   "&End Part"
         Height          =   465
         Left            =   405
         TabIndex        =   29
         Top             =   7710
         Width           =   975
      End
      Begin VB.ComboBox Bookname 
         Height          =   2640
         Left            =   4800
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   28
         Top             =   2400
         Width           =   3645
      End
      Begin VB.ComboBox Bookcode 
         Height          =   2640
         ItemData        =   "frmBookIssueSp_Ret.frx":0000
         Left            =   1440
         List            =   "frmBookIssueSp_Ret.frx":0002
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   27
         Top             =   2385
         Width           =   3210
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   360
         ScaleHeight     =   756
         ScaleWidth      =   9156
         TabIndex        =   17
         Top             =   8580
         Width           =   9150
         Begin VB.CommandButton Commandprintnh 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Enabled         =   0   'False
            Height          =   645
            Left            =   6855
            Picture         =   "frmBookIssueSp_Ret.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   45
            Width           =   1065
         End
         Begin VB.CommandButton Commandadd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   645
            Left            =   75
            Picture         =   "frmBookIssueSp_Ret.frx":0BE8
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   45
            Width           =   1065
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   645
            Left            =   7965
            Picture         =   "frmBookIssueSp_Ret.frx":17CC
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   45
            Width           =   1065
         End
         Begin VB.CommandButton CommandPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Enabled         =   0   'False
            Height          =   195
            Left            =   8700
            Picture         =   "frmBookIssueSp_Ret.frx":23B0
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   45
            Width           =   75
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   645
            Left            =   5715
            Picture         =   "frmBookIssueSp_Ret.frx":2F94
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   45
            Width           =   1065
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Enabled         =   0   'False
            Height          =   645
            Left            =   4590
            Picture         =   "frmBookIssueSp_Ret.frx":3B78
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   45
            Width           =   1065
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   645
            Left            =   3465
            Picture         =   "frmBookIssueSp_Ret.frx":475C
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   45
            Width           =   1065
         End
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Height          =   645
            Left            =   2340
            Picture         =   "frmBookIssueSp_Ret.frx":4CE6
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   45
            Width           =   1065
         End
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   645
            Left            =   1200
            Picture         =   "frmBookIssueSp_Ret.frx":58CA
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   45
            Width           =   1065
         End
         Begin VB.CommandButton Commandhelp 
            Caption         =   "Help"
            Height          =   375
            Left            =   -855
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   800
         End
      End
      Begin VB.ComboBox Genledger 
         Height          =   315
         Left            =   7275
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   7935
         Visible         =   0   'False
         Width           =   645
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5520
         Left            =   180
         TabIndex        =   13
         Top             =   1800
         Width           =   12525
         _ExtentX        =   22098
         _ExtentY        =   9737
         _Version        =   393216
         FillStyle       =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox bundles 
         Height          =   288
         Left            =   8976
         TabIndex        =   4
         Top             =   684
         Width           =   2436
         _ExtentX        =   4297
         _ExtentY        =   508
         _Version        =   393216
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
         Height          =   312
         Left            =   1296
         TabIndex        =   1
         Top             =   684
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tempmeb 
         Height          =   285
         Left            =   300
         TabIndex        =   32
         Top             =   2340
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
         Left            =   360
         TabIndex        =   33
         Top             =   5400
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
         Left            =   360
         TabIndex        =   34
         Top             =   5580
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3260
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_NO 
         Height          =   312
         Left            =   192
         TabIndex        =   0
         Top             =   684
         Width           =   1068
         _ExtentX        =   1884
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox marka 
         Height          =   288
         Left            =   6876
         TabIndex        =   3
         Top             =   684
         Width           =   2016
         _ExtentX        =   3556
         _ExtentY        =   508
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox weight 
         Height          =   312
         Left            =   10176
         TabIndex        =   12
         Top             =   1428
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox freight 
         Height          =   312
         Left            =   7620
         TabIndex        =   10
         Top             =   1428
         Width           =   1656
         _ExtentX        =   2921
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox bdated 
         Height          =   312
         Left            =   6480
         TabIndex        =   9
         Top             =   1428
         Width           =   1152
         _ExtentX        =   2032
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox biltno 
         Height          =   312
         Left            =   5256
         TabIndex        =   8
         Top             =   1428
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox station 
         Height          =   312
         Left            =   180
         TabIndex        =   6
         Top             =   1428
         Width           =   5028
         _ExtentX        =   8869
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNSCHNo 
         Height          =   312
         Left            =   11472
         TabIndex        =   5
         Top             =   684
         Width           =   948
         _ExtentX        =   1672
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777152
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NS-CH.No "
         ForeColor       =   &H80000008&
         Height          =   288
         Left            =   11472
         TabIndex        =   58
         Top             =   504
         Width           =   900
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Godown : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9390
         TabIndex        =   57
         Top             =   1140
         Width           =   645
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transport"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4500
         TabIndex        =   56
         Top             =   1140
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agent :"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   2340
         TabIndex        =   55
         Top             =   480
         Width           =   516
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F4 Key To Delete A Invoive Item"
         Height          =   225
         Left            =   4020
         TabIndex        =   54
         Top             =   9555
         Width           =   3705
      End
      Begin VB.Label labelbybanklbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "By Bank : "
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   53
         Top             =   7890
         Width           =   1200
      End
      Begin VB.Label labelbybank 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5190
         TabIndex        =   52
         Top             =   7950
         Width           =   1245
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6615
         TabIndex        =   51
         Top             =   1140
         Width           =   1185
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Marka : "
         ForeColor       =   &H80000008&
         Height          =   288
         Left            =   7128
         TabIndex        =   50
         Top             =   504
         Width           =   1608
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Quantity : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2610
         TabIndex        =   49
         Top             =   7605
         Width           =   1470
      End
      Begin VB.Label tqu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5205
         TabIndex        =   48
         Top             =   7665
         Width           =   1215
      End
      Begin VB.Label mgd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10980
         TabIndex        =   47
         Top             =   7650
         Width           =   1200
      End
      Begin VB.Label mna 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10980
         TabIndex        =   46
         Top             =   7965
         Width           =   1200
      End
      Begin VB.Label mga 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9720
         TabIndex        =   45
         Top             =   7680
         Width           =   1200
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   288
         Left            =   1296
         TabIndex        =   44
         Top             =   456
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Challan No. : "
         ForeColor       =   &H80000008&
         Height          =   288
         Left            =   216
         TabIndex        =   43
         Top             =   456
         Width           =   1068
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Net Amount : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9780
         TabIndex        =   42
         Top             =   7965
         Width           =   1200
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Gross Amount : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9720
         TabIndex        =   41
         Top             =   7410
         Width           =   1260
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bundle(s) : "
         ForeColor       =   &H80000008&
         Height          =   288
         Left            =   8976
         TabIndex        =   40
         Top             =   504
         Width           =   2436
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks : "
         ForeColor       =   &H80000008&
         Height          =   288
         Left            =   1296
         TabIndex        =   39
         Top             =   1176
         Width           =   948
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bilty No. : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5250
         TabIndex        =   38
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Freight : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7800
         TabIndex        =   37
         Top             =   1140
         Width           =   1410
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Weight : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9990
         TabIndex        =   36
         Top             =   1140
         Width           =   1545
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Total Discount : "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10980
         TabIndex        =   35
         Top             =   7410
         Width           =   1290
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   825
         Left            =   315
         Top             =   8535
         Width           =   9225
      End
   End
End
Attribute VB_Name = "frmBookIssueSp_Ret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kk As ADODB.Recordset
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
Dim emptyInv_bool As Boolean
Sub printinvoice()

Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = True
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = True
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
        Print #1, Tab(1); kkk!COURT; Tab(60); "FOR " + Trim(kkk!cname)
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

Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("SPECIMEN DELIVERY CHALLAN")))) / 2 - 3); Chr(14); "SPECIMEN DELIVERY CHALLAN"; Chr(20)
Line = Line + 1

'''If Printheader = True Then
  ''' Print #1, Tab(63); kkk!cst
   '''Line = Line + 1
'''End If


If Printheader = False Then
   Print #1, ""
   Line = Line + 1
End If

Print #1, repli("-", 96)
Line = Line + 1
If rs1.State = 1 Then rs1.close
rs1.Open "invoicea_spRet", con, adOpenStatic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); "AGENT NAME : "; Chr(27) + Chr(71); Mid$(rs1!agentname, 1, 20); Tab(45); Chr(27) + Chr(71); "  Challan No. : "; Chr(27) + Chr(71); Trim(rs1!invoiceNo); Tab(75); Chr(27) + Chr(71); "  Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!invoiceDate), "", rs1!invoiceDate)
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
      
      Print #1, Tab(45); Chr(27) + Chr(71); "Bilty NO.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(75); Chr(27) + Chr(71); "Dt  : "; Chr(27) + Chr(72); IIf(IsNull(Trim(rs1!BILTYDATE)), "", Trim(rs1!BILTYDATE))
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, Tab(0); Chr(27) + Chr(71); "(" & cboGodown & ")"; Chr(27) + Chr(72)
      Print #1, Chr(27) + Chr(71); "Station   :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!station); " "; Trim(rs1!transportname); Tab(75); Chr(27) + Chr(71); "Pvt. Mark : "; Chr(27) + Chr(72); Trim(rs1!marka)
      
       Print #1, Chr(27) + Chr(71); "Freight   :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!freight); Tab(40); Chr(27) + Chr(71); "Weight  : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."; Tab(67); Chr(27) + Chr(71); "Bundle(s)  : "; Chr(27) + Chr(72); Trim(rs1!bundles)
       Print #1, Chr(27) + Chr(71); repli("-", 96)
        Print #1, Tab(0); "S.No."; Tab(15); "Book Description"; Tab(50); "Quantity"; Tab(62); "Rate"; Tab(74); "Amount"; Tab(86); "Net Amount"
        Print #1, repli("-", 96); Chr(27) + Chr(72)
        Line = Line + 11
''''''    End If
    If called1 Then
        GoTo printagain1
    End If
    If called2 Then
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.close
    kk.Open "select * from invoiceb_spRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " order by printorder,sno ", con, adOpenStatic, adLockReadOnly, adCmdText
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
                If tdata.EOF = False Then
                   Print #1, Tab(0); rsets(Trim(Str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(Str(kk!QUANTITY)), 5); Tab(58); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                End If
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from invoiceb_spRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and printorder =" + Trim(Str(cdiscount)) + " group by printorder", con, adOpenStatic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(Str(tdata(0)), "0.00")), 12)
                   'Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(Str(tdata(0) * cdiscount / 100), "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata(0) - (tdata(0) * cdiscount / 100)), "0.00")), 12)
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
       kk.Open "Select * from invoicec_spRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenStatic, adLockReadOnly, adCmdText
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
        kk.Open "Select * from invoicea_spRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenStatic, adLockReadOnly, adCmdText
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
        Print #1, Tab(1); tempdata!COURT; Tab(60); "FOR " + Trim(tempdata!cname)
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
'frmEndPartTrans.calc
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
        Dim ctl As Control
        For Each ctl In Me.Controls
            If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
                If UCase(Trim(ctl.Name)) <> UCase(Trim("cbogodown")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) Then
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
                Set RS = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(Grid1.Text) & "'")
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
                                
                                
                             'If Edit = False Then
                                
                                Grid1.Col = 5
                                If Grid1.Text = "" Then
                                Grid1.Text = Format(RS(3), "0.00")            'rs(3)
                                End If
                                r = RS(3)
                                
                                Grid1.Col = 4
                                If Grid1.Text = "" Then
                                Grid1.Text = Format(RS(4), "0.00")
                                End If
                                
                                Grid1.Col = 6
                                If Grid1.Text = "" Then
                                Grid1.Text = Format(RS(4), "0.00")
                                End If
                                
                                D = RS(4)
                            'End If
                            
                                Grid1.Col = 7
                                Grid1.Text = Format(Round(q * r, 2), "0.00")
                                Grid1.Col = 8
                                Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
                             ' Else
                              
                                  If Grid1.Text = "" And addmode = False Then
                                    If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        kk.close
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
                                        Grid1.Col = 4
                                    '    If kk.BOF Then
                                     '        GoTo abc
                                     '   End If
                                        Grid1.Text = Format(kk(0), "0.00")
                                        Grid1.Col = 6
                                        Grid1.Text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = RS(3)
                                    End If
                                  End If
  
  '                            End If
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
                           Commandother.SetFocus
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
                RS.Open "books", con, adOpenDynamic, adLockReadOnly, adCmdTable
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


                                 Grid1.Col = 4
                                 Grid1.Text = Format(RS(4), "0.00")
                                    Grid1.Col = 6
                                    Grid1.Text = Format(RS(4), "0.00")
                                    D = RS(4)
                      '          End If
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

Private Sub cmbAgentName_LostFocus()

If cmbAgentName.Text = "" Then
   MsgBox "Enter a Agent Name.. "
   cmbAgentName.SetFocus
   cmbAgentName.Text = ".."
   Exit Sub
Else
  Dim rs1 As New ADODB.Recordset
  'rs1.Open "select *  from AgentMaster where AgentName='" & cmbAgentName.Text & "' and " & stringyear & " order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
  rs1.Open "select rep  from rep where rep='" & cmbAgentName.Text & "' order by rep", CON_blue, adOpenDynamic, adLockReadOnly, adCmdText
  If rs1.RecordCount <= 0 Then
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
Set RS = con.Execute("exec searchList 'SPRET'")

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

Private Sub CommandPrint_LostFocus()
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub

Private Sub Commandprintnh_Click()
's1 = 2
'Printheader = False
'printinvoice

s1 = 3
printch = "INVOICEASP_ret"
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
       'If CON.Execute("Select max(invoiceno) from invoicea_spRet")(0) >= Val(Trim(Me.I_NO.Text)) Then
       '   Me.I_NO.Text = CON.Execute("Select max(invoiceno) from invoicea_spRet")(0) + 1
       '   rs.Open "tempinv", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
       '   If rs.BOF Then
       '      rs.AddNew
       '   End If
       '   Me.I_NO.Text = rs!In + 1
       '   rs!In = Val(Me.I_NO.Text)
       '   rs.update
       '   rs.close
       
       RS.Open "SELECT MAX(INVOICENO) FROM invoicea_spRet", con, adOpenDynamic, adLockReadOnly, adCmdText
       If kk(0) <> "" Then
          Me.I_NO.Text = Trim(Str(RS(0) + 1))
       Else
          Me.I_NO.Text = "1"
       End If
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
    ''Me.customercode.Enabled = True
    cboGodown.ListIndex = 0
    'i_dt.SetFocus
    I_NO.SetFocus
    
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
End Sub
Private Sub Commandall_Click()
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Dim myvalue As String

'''If Trim(Me.customercode.Text) = "" Then
'''    MsgBox "Please Fill the customer detail "
'''    Exit Sub
'''End If

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
            Grid1.Text = RS(0)
            Grid1.Col = 2
            Grid1.Text = RS(1)
            Grid1.Col = 3
            If Trim(Grid1.Text) = "" Then
                Grid1.Text = Val(myvalue)
            End If
            q = Val(Grid1.Text)
            Grid1.Col = 5
            Grid1.Text = Format(RS(3), "0.00")            'rs(3)
            r = RS(3)
            '/******************
            Set kk = con.Execute("select DISCATEGORY from sledger where " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
            Grid1.Col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.close
                Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
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
            totalamount = totalamount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.Col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.Col = 3
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
     Next
     maxrow = Grid1.rows - 1
Else
'Grid1_Click
Exit Sub
End If

invoicecalc

End Sub

Private Sub Commanddelete_Click()

    
    
    
    
 '=====================================================================================

    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
   
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from invoicea_spRet where " & stringyear & " and invoiceno=" & I_NO.Text & "", con
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select * from invoicea_spRet where " & stringyear & " and invoiceno=" & I_NO.Text & "", con
       'If rs_h.Fields("Print_yes").Value = "y" Then
          If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
       '   End If
       End If
       
    End If
    
'======================================================================================
    
    
'===================================================================================
    
 
  
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
      

If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
                con.Execute ("delete from invoicea_spRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete from invoiceb_spRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete from invoicec_spRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                invoiceabandon
End If
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
    'Me.customercode.Enabled = True
    Edit = True
    addoredit = False
    'I_NO_LostFocus
    i_dt.Enabled = True
    i_dt.SetFocus
    
    ' INVOICECtmp_spRet creation start
    con.Execute ("delete  from INVOICECtmp_spRet WHERE " & stringyear & " and username='" & UserName & "' and INVOICENO = " + Trim(I_NO.Text))
    DoEvents
    
    
      con.Execute ("insert into INVOICECtmp_spRet([INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],[UserName]) " & _
    "  select [INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],'" & UserName & "' " & _
    " from invoicec_spRet where  " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
    '=================================
  
    
    
     Dim ctl As Control
    For Each ctl In Me.Controls
    
    If (TypeOf ctl Is Label Or TypeOf ctl Is textbox Or TypeOf ctl Is MaskEdBox Or TypeOf ctl Is ComboBox) Then
        ctl.Enabled = True
    End If
    
    Next
 
    
    Grid1.Enabled = True
    i_dt.Enabled = True
    i_dt.SetFocus
    
    
    addoredit = False
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
End Sub
Private Sub Commandother_Click()

Commandsave.Enabled = True
searchForm = "invoice_spret"
frmEndPartTrans.Show
  
 
End Sub
Private Sub CommandPrint_Click()
  Printheader = True
  printinvoice
   
End Sub

Private Sub Commandprintnh_LostFocus()
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub

Private Sub CommandReturn_Click()
 
   
   Unload Me
   addoredit = False
    
End Sub
Private Sub Commandsave_Click()
    
 On Error GoTo aa1
    
'=====================================================================================

    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
   
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from invoicea_spRet where " & stringyear & " and invoiceno=" & I_NO.Text & "", con
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select * from invoicea_spRet where " & stringyear & " and invoiceno=" & I_NO.Text & "", con
          If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
       End If
       
    End If
    
'======================================================================================
    
    
    
    Dim SAVED As Boolean
    Dim LAMOUNT As Double
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
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
    
     If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(Me.cmbAgentName.Text) <> "" Then
            If Edit Then
                
                con.BeginTrans
                con.Execute ("delete  from invoicea_spRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete  from invoiceb_spRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete  from invoicec_spRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete  from cashregister where CMNo = " + Trim(I_NO.Text))
                con.Execute ("delete  from INVOICEB_spRet_Free where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                
            End If
            If RS.State = 1 Then
                RS.close
            End If
            LAMOUNT = 0
            RS.Open "select * from invoicea_spRet where " & stringyear & " and invoiceno <=0", con, adOpenDynamic, adLockOptimistic
            If Not Edit Then
again:
           If con.Execute("Select max(invoiceno) from invoicea_spRet")(0) >= Val(Trim(Me.I_NO.Text)) Then
                   ' Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'GoTo again
                End If
            End If
            RS.AddNew
            RS!invoiceNo = Val(Me.I_NO.Text)
            RS!invoiceDate = Me.i_dt.Text
            RS!Godown = cboGodown.Text
            RS!agentname = Trim(Me.cmbAgentName.Text)
            RS!transportname = Trim(Me.cmbtransportname.Text)
            
            RS!marka = Trim(Me.marka.Text)
            RS!bundles = Trim(Me.bundles)
            RS!NsChallanNo = Trim(txtNSCHNo.Text)
            
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
            
err1:
           If Not Edit Then
                If con.Execute("Select max(invoiceno) from invoicea_spRet")(0) >= Val(Trim(Me.I_NO.Text)) Then
                    On Error GoTo err1
                End If
            End If
            
            RS!setupid = setupid
            RS!fyear = session
            
            If kk.State = 1 Then kk.close
            kk.Open "select Add1,Add2,City,pin,District,State,Email from SalesRepQry where rep='" & Trim(Me.cmbAgentName.Text) & "'", CON_blue
            If kk.EOF = False Then
              RS!add1 = kk!add1
              RS!add2 = kk!add2
              RS!city = kk!city & "-" & kk!pin
              RS!District = kk!District
              RS!states = kk!State
              RS!mail = kk!email
            End If

            
            
            RS.update
            On Error GoTo 0
            RS.close
            RS.Open "select * from invoiceb_spRet where " & stringyear & " and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
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
                         RS!invoiceNo = Val(Me.I_NO.Text)
                         RS!invoiceDate = Me.i_dt.Text
             '           rs!Genledger = Trim(Me.Genledger.Text)
             '           rs!SUBLEDGER = Trim(Me.customercode.Text)
                         RS!Bookcode = Trim(Grid1.Text)
                         Grid1.Col = 3
                         RS!QUANTITY = Trim(Grid1.Text)
                         Grid1.Col = 5
                         RS!rate = Trim(Grid1.Text)
                         Grid1.Col = 7
                         RS!amount = Trim(Grid1.Text)
                         LAMOUNT = Val(Trim(Grid1.Text))
                         Grid1.Col = 4
                         RS!PRINTORDER = Trim(Grid1.Text)
                         Grid1.Col = 6
                         RS!discount = Trim(Grid1.Text)
                         Grid1.Col = 8
                         RS!netamount = LAMOUNT - Trim(Grid1.Text)
                         LAMOUNT = 0
                         RS!agentname = Trim(Me.cmbAgentName.Text)
                         RS!BookDesc = ReturnBookDesc(Grid1.TextMatrix(I, 0))
                         RS!setupid = setupid
                         RS!fyear = session
                         
                                                 
                         If kk.State = 1 Then kk.close
                         kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,apply from KitQry where kitcode='" & RS!Bookcode & "'", con
                         While kk.EOF = False
                         
                         If kk!Apply = "y" Then
                           con.Execute "insert into INVOICEB_spRet_Free(INVOICENO,INVOICEDATE,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
                           " values('" & Val(Me.I_NO.Text) & "','" & Format(Me.i_dt.Text, "MM/dd/yyyy") & "','" & kk!Bookcode & "','" & (kk!qty * RS!QUANTITY) & "','" & kk!rate & "','" & Trim(Me.cmbAgentName.Text) & "','" & setupid & "','" & session & "','" & cboGodown & "')"
                         End If
                         
                         kk.MoveNext
                         
                         Wend
                         
                         
                         RS.update
                       End If
                    End If
                End If
            Next
            RS.close
            Grid1.TopRow = 1
            Grid1.Row = 1
            Grid1.Col = 1
            RS.Open "select * from invoicec_spRet where " & stringyear & " and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
            '/******
                'Dim I, x As Integer
                Dim temprs As ADODB.Recordset
                Set temprs = New ADODB.Recordset
                For I = 1 To frmEndPartTrans.vs.rows - 1
                    frmEndPartTrans.vs.Row = I
                    'frmEndPartTrans.vs.Col = 0
                    If Trim(frmEndPartTrans.vs.TextMatrix(I, 0)) <> "" Then
                        RS.AddNew
                        RS!invoiceNo = Val(Me.I_NO.Text)
                        RS!invoiceDate = Me.i_dt.Text
                        RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
                        RS!Text = Trim(frmEndPartTrans.vs.TextMatrix(I, 0))            'Trim(frmEndPartTrans.vs.Text)
                        If temprs.State = 1 Then
                            temprs.close
                        End If
                         
                        If Edit Then
                        temprs.Open "select * from INVOICECtmp_spRet where " & stringyear & " and INVOICENO=" & frmBookIssueSp_Ret.I_NO & "", con, adOpenDynamic, adLockReadOnly, adCmdText
                        If Trim(frmEndPartTrans.vs.TextMatrix(I, 0)) <> "" Then
                                'temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.Text) + "'"
                                temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.TextMatrix(I, 0)) + "'"
                                RS!DebitorCredit = Trim(temprs!DebitorCredit)
                                RS!RYN = temprs!RYN & ""
                                
                        End If
                        temprs.close
                        
                        
                        Else
                        
                        temprs.Open "select * from INVOICEEND  order by printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
                        If Trim(frmEndPartTrans.vs.TextMatrix(I, 0)) <> "" Then
                                temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.TextMatrix(I, 0)) + "'"
                                RS!DebitorCredit = Trim(temprs!DebitorCredit)
                                RS!RYN = temprs!RYN & ""
                        End If
                        temprs.close
                        End If
                        'frmEndPartTrans.vs.Col = 1
                        RS!rate = Val(Trim(frmEndPartTrans.vs.TextMatrix(I, 1)))
                        If Val(Trim(frmEndPartTrans.vs.TextMatrix(I, 1))) > 0 Then
                            RS!amount = Round((Me.totalamount - Me.totaldiscount), 2) * Round((Val(Trim(frmEndPartTrans.vs.TextMatrix(I, 1))) / 100), 2)
                        Else
                        'frmEndPartTrans.vs.Col = 2
                        RS!amount = Val(Trim(frmEndPartTrans.vs.TextMatrix(I, 2)))
                        End If
                    RS!setupid = setupid
                    RS!fyear = session
                    RS!UserName = UserName
                    
                    RS.update
                    End If
                Next
                RS.close
              
              'CON.Execute ("delete  from INVOICECtmp_spRet where  " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
              If Me.station.Text <> "" Then
                    
                    s11 = ""
                    ss11 = ""
                    
                    s11 = InStr(1, Me.station.Text, " ")
                    If s11 <> 0 Then
                    ss11 = Trim(Mid(Me.station.Text, 1, s11))
                    Else
                    ss11 = Me.station.Text
                    End If
                    PopUpValue1 = ss11

                 UpdateDisPatchReg1 I_NO, i_dt, Me.cmbAgentName.Text, PopUpValue1, Trim(Me.bundles), Trim(Me.cmbtransportname.Text), "-", Trim(Me.biltno.Text), Me.bdated, Trim(Me.freight), "CashRegister"
                 PopUpValue1 = ""
              End If

                
                
                
            SAVED = True
            
           If Edit Then
              con.CommitTrans
           End If
            

            
            
        End If
        If SAVED Then
            MsgBox "Record Saved"
            Unload frmEndPartTrans
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
aa1:
Set RS = New ADODB.Recordset

If err.Number = -2147217887 Then
    con.RollbackTrans
End If

MsgBox err.DESCRIPTION
     
        
End Sub

Private Sub Commandsearch_Click()
   searchType = "inv"
   
   sqlQry = "select InvoiceNo,InvoiceDate,Subledger,NetAmount from invoicea_spRet where InvoiceNo"
   orderby = "order by InvoiceNo"

   
   popuplist10 "select InvoiceNo,InvoiceDate,AgentName as Representative,NetAmount from invoicea_spRet where " & stringyear & "  order by InvoiceNo", con

End Sub

Private Sub customercode_LostFocus()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "select * from sledger where " & stringyear & " and gledger='SUNDRY DEBTORS' and subledger='" + Trim(customercode.Text) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
    If RS.RecordCount <= 0 Then
        customercode.SetFocus
        HIT
        RS.close
        Exit Sub
    End If
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    If RS!distcode <> "" And addmode = True Then
       rs1.Open "Select * from Districts  where " & stringyear & " and Districtname = '" & RS!distcode & "'", con, adOpenStatic, adLockReadOnly
       If rs1.RecordCount > 0 Then
          Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
       End If
    End If
    RS.close
  
  
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

Private Sub Commandsearch_GotFocus()
If PopUpValue1 <> "" Then
     I_NO.Text = PopUpValue1
     I_NO_LostFocus
     i_dt.SetFocus
     PopUpValue1 = ""
End If
End Sub

Private Sub Form_Activate()
'SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
   Unload Me
End If


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
  
Me.top = 100
Me.Left = 100

'Me.Width = 11500
'Me.Height = 8500

Me.Width = 13140
Me.Height = 10455



  
  
Screen.MousePointer = vbHourglass
  
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
    Grid1.rows = 2
    Grid1.Cols = 1
    Grid1.rows = 10
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
    Grid1.ColWidth(0) = 200
    Grid1.ColWidth(1) = 1100
    Grid1.ColWidth(2) = 3800
    Grid1.ColWidth(3) = 1050
    Grid1.ColWidth(4) = 1050
    Grid1.ColWidth(5) = 1050
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1400
    Grid1.ColWidth(8) = 1400
    Bookname.Height = 2325
    Me.CommandPrint.Enabled = True
    Me.Commandprintnh.Enabled = True
    RS.Open "select * from books", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            Me.Bookcode.AddItem RS(0)
            Me.Bookname.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    Genledger.Text = "SUNDRY DEBTORS"
    
     
''    '*******Agent  combo fill
''    RS.Open "select  Agentname from AgentMaster where " & stringyear & "  order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
''    cmbAgentName.Clear
''    If Not RS.EOF Then
''       Do While Not RS.EOF
''          If IsNull(RS(0)) = False Then
''            Me.cmbAgentName.AddItem RS(0)
''          End If
''          If Not RS.EOF Then RS.MoveNext
''        Loop
''    End If
''    RS.Close

'*******Agent  combo fill
If RS.State = 1 Then RS.close
'RS.Open "select  Agentname from AgentMaster where " & stringyear & " order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
RS.Open "select  rep from rep where (email is not null and len(email)>1) order by rep", CON_blue, adOpenDynamic, adLockReadOnly, adCmdText
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
    
    
    
    
    RS.Open "select transportname from transportMaster order by transportname", con, adOpenDynamic, adLockReadOnly, adCmdText
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

    
    
    
    RS.Open "select Godwn from godownMaster where Binder_Printer='g' order by id"
    While RS.EOF = False
        'If Len(RS(0)) = 1 Then
          cboGodown.AddItem RS(0)
        'End If
          RS.MoveNext
    Wend
    RS.close
    
    On Error Resume Next
    
    Bookcode.Left = Grid1.Left
    Bookcode.Visible = False
    Bookname.Visible = False
    Grid1.rows = 100
    For I = 1 To 99
        Grid1.RowHeight(I) = 300
    Next
    Bookcode.Width = 1230
    Bookname.Width = 2830
    amount.Width = rate.Width
  
'''    rs.Open "tempinv", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
'''    If Not rs.BOF Then
'''        Me.I_NO.Text = rs!In
'''        INVOICE.Enabled = True
'''        INVOICE.Edit = False
'''        INVOICE.I_NO_LostFocus
'''        INVOICE.I_NO.Enabled = False
'''        lastrow = 0
'''        lastcol = 1
'''        Dim ctl As Control
'''        For Each ctl In INVOICE.Controls
'''            If Not TypeOf ctl Is CommandButton Then
'''                ctl.Enabled = False
'''            End If
'''            If UCase(Trim(ctl.Name)) = UCase(Trim(INVOICE.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(INVOICE.Commandall.Name)) Then
'''                ctl.Enabled = False
'''            End If
'''        Next
        
'        INVOICE.Picture5.Enabled = True
'        addoredit = False
'        SendKeys "{TAB}"
'    Else
       kk.Open "SELECT MAX(INVOICENO) FROM invoicea_spRet", con, adOpenDynamic, adLockReadOnly, adCmdText
       If kk(0) <> "" Then
          Me.I_NO.Text = Trim(Str(kk(0)))
          I_NO_LostFocus
       Else
          Me.I_NO.Text = "1"
       End If
       kk.close
'    End If
'    rs.close
    
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
    'cboGodown.ListIndex = 0
        
        
    BackColorFrom Me
    
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    
    Commandsave.Enabled = False
    
Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Resize()

panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2
End Sub

Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub

Private Sub Grid1_Click()

If Trim(Me.cmbAgentName.Text) <> "" Then


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
        
            
If Trim(Me.cmbAgentName.Text) <> "" Then
            If Me.cmbAgentName.Enabled = True Then
                Me.cmbAgentName.Enabled = False
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
    
        
        If Trim(Me.cmbAgentName.Text) <> "" Then
        If Me.cmbAgentName.Enabled = True Then
            Me.cmbAgentName.Enabled = False
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
'If Trim(Me.customercode.Text) <> "" Then
If Trim(Me.cmbAgentName.Text) <> "" Then

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
                tempmeb.Left = Grid1.CellLeft + Grid1.Left
                tempmeb.top = Grid1.top + Grid1.CellTop '- 50
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.Text = Grid1.Text
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
   PopupMenu dd, , Grid1.Left + X, Grid1.top + Y
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

Private Sub i_dt_LostFocus()
On Error Resume Next
If Trim(i_dt.Text) <> Trim("__/__/____") Then
    If Not checkdate(Trim(i_dt.Text), i_dt) Then
        i_dt.SetFocus
    End If
    Dim tRS1 As New ADODB.Recordset
    Dim trs2 As New ADODB.Recordset
    If trs2.State = 1 Then trs2.close
    trs2.Open "Select invoiceno as cn from invoicea_spRet", con, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount <= 0 Then
       Exit Sub
    Else
'''''''        If tRS1.State = 1 Then tRS1.close
'''''''        tRS1.Open "Select min(invoiceno) as mid,invoicedate from  invoicea_spRet  group by invoiceno,invoiceDate", CON, adOpenDynamic, adLockOptimistic
'''''''
'''''''            If tRS1.RecordCount > 0 Then
'''''''            If CDate(i_dt) <= tRS1!invoicedate Then
'''''''
'''''''                      If CDate(i_dt) <> tRS1!invoicedate Then
            '''''''                 MsgBox "Please select valid Invoice No. for this date.."
'''''''                         Me.i_dt.SetFocus
'''''''                         Exit Sub
'''''''                       Else
'''''''                         If Val(I_NO) <= tRS1!Mid Then
'''''''                 ''''''''''''''MsgBox "Please select valid Invoice No. for this date.."
'''''''                 I_NO.SetFocus
'''''''                 Exit Sub
'''''''               End If
'''''''               End If
'''''''            End If
'''''''        End If
    End If
    
    If trs2.State = 1 Then trs2.close
    trs2.Open "Select max(invoiceno) as mid from invoicea_spRet where " & stringyear & " and  invoicedate <= cdate('" & i_dt.Text & "')-1", con, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount > 0 Then
    
        If IsNull(trs2!Mid) <> True Then
            If Val(I_NO.Text) >= trs2!Mid Then
               If tRS1.State = 1 Then tRS1.close
               tRS1.Open "Select  min(InvoiceNo)as m2 from invoicea_spRet where " & stringyear & " and invoicedate >= cdate('" & i_dt.Text & "')+1", con, adOpenDynamic, adLockOptimistic
               If tRS1.RecordCount > 0 Then
                  If IsNull(tRS1!m2) <> True Then
                     If Val(I_NO.Text) <= tRS1!m2 Then
                       
                     Else
                         MsgBox "Please select valid Invoice No for this date.."
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


Private Sub I_DTOB_LostFocus()
'If Trim(I_DTOB.Text) <> "__/__/____" Then
'    If Not checkdate(Trim(I_DTOB.Text), I_DTOB) Then
'        I_DTOB.SetFocus
'    End If
'End If
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
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
    
    
    If Val(inviceNo) > 0 Then
    I_NO.Text = inviceNo
      ' cmdButto
    End If
    
    inviceNo = ""

    
    
    
    If Trim(I_NO.Text) = "" Then
        MsgBox "Invoice cannot be null"
        I_NO.SetFocus
    Else
        If RS.State = 1 Then
           RS.close
        End If
        ''rs.Open "invoicea_spRet", con, adOpenDynamic, adLockReadOnly, adCmdTable
        RS.Open "Select * from  invoicea_spRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text) + "", con, adOpenStatic, adLockReadOnly
        ''rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
        If RS.EOF Then
            If addoredit = False Then
                MsgBox "Invoice not found"
                Exit Sub
            End If
            Exit Sub
        End If
        If addoredit Then
            X = MsgBox("Invoice already exist...", vbOKOnly)
            I_NO.SetFocus
            'I_NO.SetFocus
            HIT
            Exit Sub
        End If
        
        invoiceabandon
        
        Dim ctl As Control
        For Each ctl In Me.Controls
            If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = True
            End If
        Next
        
        Commanddelete.Enabled = False
        
        I_NO.Text = RS!invoiceNo
        Me.i_dt.Text = RS!invoiceDate
        cboGodown.Text = RS!Godown & ""
        Me.Genledger.Text = Trim(RS!Genledger)
        ''Me.customercode.Text = Trim(rs!SUBLEDGER)
        Me.cmbAgentName.Text = IIf(IsNull(RS!agentname), "", RS!agentname)
        Me.cmbtransportname.Text = IIf(IsNull(RS!transportname), "", RS!transportname)
        txtNSCHNo.Text = RS!NsChallanNo & ""
        
        ''Me.textbox.Text = Trim(rs!SUBLEDGER)
        '''Me.I_OB.Text = IIf(IsNull(rs!orderby), "", Trim(rs!orderby))
        ''If rs!ORDERDATE <> "" Then
        ''Me.I_DTOB.Text = rs!ORDERDATE
        ''End If
        Me.marka.Text = IIf(IsNull(RS!marka), "", Trim(RS!marka))
        Me.bundles = IIf(IsNull(RS!bundles), "", RS!bundles)
        ''Me.through.Text = IIf(IsNull(rs!through), "", rs!through)
        ''Me.through1.Text = IIf(IsNull(rs!through1), "", rs!through1)
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
       
       ' frmEndPartTrans.Form_Load
'*/**/*/*/*/*//*/*
        If RS.State = 1 Then
                RS.close
        End If
'Commandedit.Enabled = True

       ' Unload frmEndPartTrans
        con.Execute "select * from INVOICECtmp_spRet where " & stringyear & " and INVOICENO=" & frmBookIssueSp_Ret.I_NO & ""
       ' frmEndPartTrans.Form_Load
        RS.Open "Select * from invoiceb_spRet where " & stringyear & " and INVOICENO =" + Trim(I_NO.Text) + " order by SNO", con, adOpenStatic, adLockReadOnly
       '' rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
       Grid1.TopRow = 2
        If Not RS.EOF Then
        
            Grid1.Row = 1
            Grid1.Col = 1
            Do While Not RS.EOF
               If Trim(RS!invoiceNo) = Trim(I_NO.Text) Then
                Grid1.Col = 1
                Grid1.Text = Trim(RS!Bookcode)
                If kk.State = 1 Then
                    kk.close
                End If
                kk.Open "select * from books where " & stringyear & " and bookcode='" + Trim(RS!Bookcode) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
                Grid1.Col = 2
                Grid1.Text = Trim(kk!Bookname)
                Grid1.Col = 3
                Grid1.Text = Trim(RS!QUANTITY)
                Grid1.Col = 5
                Grid1.Text = Format(Round(RS!rate, 2), "0.00")
                Grid1.Col = 7
                Grid1.Text = Format(Round(RS!amount, 2), "0.00")
                Grid1.Col = 4
                
                Grid1.Text = Format(Round(RS!PRINTORDER, 2), "0.00")
                Grid1.Col = 6
                
                Grid1.Text = Format(Round(RS!discount, 2), "0.00")
                Grid1.Col = 8
                Grid1.Text = Format(Round(RS!amount * (RS!discount / 100), 2), "0.00")
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
       'templost = True
    End If
    Me.Commandother.Enabled = True
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
If KeyAscii = 13 Then
        Dim RS As ADODB.Recordset
           Set RS = New ADODB.Recordset
            Select Case Grid1.Col
                Case 1
                    'If RS.State = 1 Then
                    '    RS.close
                    'End If
                    'RS.Open "books", CON, adOpenStatic, adLockReadOnly, adCmdTable
                    
                    Set RS = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(Grid1.Text) & "'")
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
    'Me.customercode.Enabled = True
   ' Me.customercode.Visible = True
  '  Me.customercode.Height = 1100
'Me.customercode.ZOrder
 '   Me.customercode.SetFocus
    
End Sub
Private Sub through_LostFocus()
through = UCase(through)
End Sub
Private Sub through1_LostFocus()
through1 = UCase(through1)
End Sub

Private Sub weight_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(Me.cmbAgentName.Text) <> "" Then
            Me.Grid1.SetFocus
            Grid1.Col = 1
            Grid1.Row = 1
            Grid1_Click
        Else
            Me.cmbAgentName.SetFocus
            'Me.customercode.SetFocus
        End If
    End If
'''If KeyAscii = 13 Then
'''        If Trim(Me.customercode.Text) <> "" Then
'''            Grid1.col = 1
'''            Grid1.row = 1
'''            Grid1_Click
'''        Else
'''            Me.textbox.SetFocus
'''            'Me.customercode.SetFocus
'''        End If
'''    End If
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
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
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
          kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
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
    rs1.Open "invoicea_spRet", con, adOpenDynamic, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
   
    Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!subledger; Tab(T5); "Invoice No. : "; Trim(rs1!invoiceNo); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!invoiceDate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.close
            End If
            kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
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
            kk.Open "select * from invoiceb_spRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " order by discount,printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                        tdata.Open "Select bookname from books where " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                        tdata.Open "select sum(amount) from invoiceb_spRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
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
           kk.Open "Select * from invoicec_spRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
           kk.Open "Select * from invoicea_spRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
Me.Commandedit.Enabled = True
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = True
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
   rs1.close
End If
'Print #1, Chr(27) + Chr(14)
'line = line + 1
If rs1.State = 1 Then
    rs1.close
End If
rs1.Open "invoicea_spRet", con, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To, S.L. Code : "; Tab(T1 + 10); Mid$(rs1!subledger, 1, 5); Tab(T5); "Invoice No. : "; Trim(rs1!invoiceNo); Tab(T8); "Dated     : "; rs1!invoiceDate   'Chr(27) + Chr(18);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
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
    kk.Open "select * from invoiceb_spRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " order by discount,printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "select sum(amount) from invoiceb_spRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
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
        kk.Open "Select * from invoicec_spRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
        kk.Open "Select * from invoicea_spRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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




