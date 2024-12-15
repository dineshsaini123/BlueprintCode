VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form countersale 
   ClientHeight    =   7830
   ClientLeft      =   270
   ClientTop       =   1815
   ClientWidth     =   10785
   Icon            =   "cashsale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboCatII1 
      Height          =   315
      Left            =   9180
      TabIndex        =   8
      Top             =   495
      Width           =   915
   End
   Begin VB.ComboBox cboCatII 
      Height          =   315
      Left            =   7845
      TabIndex        =   7
      Top             =   495
      Width           =   1050
   End
   Begin VB.ComboBox txtMark 
      Height          =   315
      ItemData        =   "cashsale.frx":000C
      Left            =   7260
      List            =   "cashsale.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1410
      Width           =   930
   End
   Begin VB.ComboBox cmbtransportname 
      Height          =   315
      Left            =   1380
      TabIndex        =   15
      Top             =   1440
      Width           =   1965
   End
   Begin VB.ComboBox cmbareaname 
      Height          =   1155
      Left            =   6300
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   11
      Top             =   105
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.ComboBox cmbAgentName 
      Height          =   315
      Left            =   3420
      TabIndex        =   27
      Top             =   6300
      Width           =   2745
   End
   Begin VB.ComboBox cmbdiscountcat 
      Height          =   315
      Left            =   6300
      TabIndex        =   6
      Top             =   480
      Width           =   1260
   End
   Begin VB.ComboBox Combosldistrictcode 
      Height          =   315
      Left            =   6300
      TabIndex        =   9
      Top             =   810
      Width           =   1275
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   60
      TabIndex        =   57
      Top             =   75
      Width           =   1740
      Begin VB.OptionButton Optioncredit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   300
         Width           =   840
      End
      Begin VB.OptionButton Optioncash 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   750
      End
   End
   Begin MSMask.MaskEdBox textbox 
      Height          =   315
      Left            =   6300
      TabIndex        =   5
      Top             =   105
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Commandall 
      Caption         =   "All Books"
      Height          =   510
      Left            =   1230
      TabIndex        =   30
      Top             =   6450
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Commandother 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&End Part"
      Height          =   510
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6450
      Width           =   930
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4185
      Left            =   135
      TabIndex        =   26
      Top             =   1770
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7382
      _Version        =   393216
      BackColorFixed  =   11206655
      ForeColorFixed  =   4210752
      GridColorFixed  =   12648447
      FillStyle       =   1
      Appearance      =   0
   End
   Begin VB.ComboBox Bookname 
      Height          =   960
      Left            =   3480
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   22
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox Bookcode 
      Height          =   765
      Left            =   420
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   21
      Top             =   2640
      Width           =   2355
   End
   Begin VB.PictureBox Picture5 
      Height          =   570
      Left            =   135
      ScaleHeight     =   510
      ScaleWidth      =   9615
      TabIndex        =   24
      Top             =   7035
      Width           =   9675
      Begin VB.CommandButton Commandprintnh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N&HPrint"
         Enabled         =   0   'False
         Height          =   510
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton Commandadd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   510
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CommandReturn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Return"
         Height          =   510
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton CommandPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   510
         Left            =   7380
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton Commandsearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
         Height          =   510
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton Commanddelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   510
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Commandabandon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aba&ndon"
         Height          =   510
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton Commandsave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sa&ve"
         Height          =   510
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Commandedit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   510
         Left            =   1110
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Commandhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   -780
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   800
      End
   End
   Begin MSMask.MaskEdBox I_DTOB 
      Height          =   315
      Left            =   3210
      TabIndex        =   13
      Top             =   735
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bdated 
      Height          =   315
      Left            =   4710
      TabIndex        =   17
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bundles 
      Height          =   285
      Left            =   8220
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox I_OB 
      Height          =   285
      Left            =   975
      TabIndex        =   12
      Top             =   780
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
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
      Height          =   285
      Left            =   3195
      TabIndex        =   4
      Top             =   420
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tempmeb 
      Height          =   285
      Left            =   480
      TabIndex        =   23
      Top             =   2130
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
      Left            =   510
      TabIndex        =   25
      Top             =   3270
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
      Left            =   600
      TabIndex        =   47
      Top             =   2730
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox I_NO 
      Height          =   285
      Left            =   3195
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.ComboBox Genledger 
      Height          =   315
      Left            =   9720
      Sorted          =   -1  'True
      TabIndex        =   48
      Top             =   1890
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.ComboBox customercode 
      Height          =   1155
      Left            =   6300
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   10
      Top             =   105
      Visible         =   0   'False
      Width           =   3795
   End
   Begin MSMask.MaskEdBox freight 
      Height          =   315
      Left            =   5910
      TabIndex        =   18
      Top             =   1410
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox biltno 
      Height          =   315
      Left            =   3330
      TabIndex        =   16
      Top             =   1440
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox station 
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.Label lbldis 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(III)"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   8925
      TabIndex        =   68
      Top             =   495
      Width           =   315
   End
   Begin VB.Label lbldis 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(II)"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   7590
      TabIndex        =   67
      Top             =   510
      Width           =   315
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mark"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7260
      TabIndex        =   66
      Top             =   1170
      Width           =   915
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Freight : "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5910
      TabIndex        =   65
      Top             =   1140
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bilty No. : "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3330
      TabIndex        =   64
      Top             =   1140
      Width           =   1365
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Railway/Station : "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   63
      Top             =   1140
      Width           =   1230
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4680
      TabIndex        =   62
      Top             =   1140
      Width           =   1230
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transport"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1380
      TabIndex        =   61
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agent :"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2295
      TabIndex        =   60
      Top             =   6300
      Width           =   1125
   End
   Begin VB.Label lbldis 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Discount Category (I)"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   4740
      TabIndex        =   59
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "District Name"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4725
      TabIndex        =   58
      Top             =   810
      Width           =   1530
   End
   Begin VB.Label labelbybanklbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By Cash : "
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2310
      TabIndex        =   56
      Top             =   6615
      Width           =   1110
   End
   Begin VB.Label labelbybank 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   300
      Left            =   3450
      TabIndex        =   55
      Top             =   6615
      Width           =   1200
   End
   Begin VB.Label mgd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7980
      TabIndex        =   51
      Top             =   6030
      Width           =   1230
   End
   Begin VB.Label mna 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8010
      TabIndex        =   50
      Top             =   6330
      Width           =   1200
   End
   Begin VB.Label mga 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6705
      TabIndex        =   49
      Top             =   6030
      Width           =   1200
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1875
      TabIndex        =   46
      Top             =   420
      Width           =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cash Memo No. : "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1875
      TabIndex        =   44
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cust Code : "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4740
      TabIndex        =   42
      Top             =   150
      Width           =   1515
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Net Amount : "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   39
      Top             =   6330
      Width           =   1200
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Gross Amount : "
      Height          =   255
      Left            =   6660
      TabIndex        =   37
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Order By : "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   135
      TabIndex        =   35
      Top             =   780
      Width           =   840
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1995
      TabIndex        =   33
      Top             =   780
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bundle(s):"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8220
      TabIndex        =   28
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Discount : "
      Height          =   255
      Left            =   7890
      TabIndex        =   52
      Top             =   2970
      Width           =   1290
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Quantity : "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2295
      TabIndex        =   54
      Top             =   5985
      Width           =   1110
   End
   Begin VB.Label tqu 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   300
      Left            =   3435
      TabIndex        =   53
      Top             =   5985
      Width           =   1245
   End
   Begin VB.Menu dd 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "countersale"
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
Public edit As Boolean
Dim addmode As Boolean
Dim Printheader As Boolean
Dim addoredit As Boolean
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
If Me.txtMark.Text = "M" Then
mystr1 = "MOHKAMPUR"
ElseIf Me.txtMark.Text = "W" Then
mystr1 = "W.K.ROAD"
ElseIf Me.txtMark.Text = "U" Then
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
    kkk.Open "select * from setup1", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
        Print #1, Tab(1); kkk!COURT; Tab(50); "FOR " + Trim(kkk!CNAME)
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
     'Print #1, ""
     'Print #1, ""
   
     Print #1, Chr(27) + Chr(71); Chr(27) + Chr(77) + Chr(14)
     Print #1, Tab((paperWidth - (Len(Trim(kkk!CNAME)) * 2)) / 2); Chr(27) + Chr(77) + Chr(14); Trim(kkk!CNAME)
     Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(77); dspace(Trim(kkk!add1))
     Print #1, Tab((paperWidth - (Len(Trim(kkk!phone1)) * 2)) / 2); Trim(kkk!phone1) & "," & Trim(kkk!phone2)
     Line = Line + 7
   End If
Else
     'Print #1, ""
     'Print #1, ""
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
rs1.Open "casha", CON, adOpenStatic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then

Print #1, Chr(27) + Chr(71); "To, S.L. Code :"; Tab(19); IIf(Optioncash.value = True, "", Mid$(rs1!SUBLEDGER, 1, 5)); Tab(38); "Cash Memo No.: "; Trim(rs1!INVOICENO); Tab(67); "Dt. : "; rs1!invoicedate; Chr(27) + Chr(72);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & IIf(Optioncash.value = True, rs1!CASHPARTYNAME, kkk!DESCFORINVOICE)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS1), "", kkk!ADDRESS1); Tab(37); Chr(27) + Chr(71); "Bilty No     : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(68); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS2), "", kkk!ADDRESS2); Tab(37); Chr(27) + Chr(71); "Bundle(s)    : "; Chr(27) + Chr(72); Trim(rs1!bundles); Tab(64); Chr(27) + Chr(71); "Freight :"; Chr(27) + Chr(72); rs1!freight
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS3), "", kkk!ADDRESS3); ; Tab(37); Chr(27) + Chr(71); "Agent Name   : "; Chr(27) + Chr(72); Trim(rs1!agentname)
        Print #1, Tab(5); "Station : " + IIf(IsNull(rs1!station), "", rs1!station) + " " + IIf(IsNull(rs1!transportname), "", rs1!transportname); Tab(73); Chr(27) + Chr(71); "(" & txtMark & ")"; Chr(27) + Chr(72)
        kkk.close
        'Print #1, Chr(27) + Chr(71); "Through  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!through) + IIf(Trim(rs1!through1) = "", "", "," & rs1!through1)
        'Print #1, Chr(27) + Chr(71); "Station  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!station); Tab(56); Chr(27) + Chr(71); "Pvt. Mark   : "; Chr(27) + Chr(72); Trim(rs1!marka)
        'Print #1, Chr(27) + Chr(71); "Freight  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!freight); Tab(35); Chr(27) + Chr(71); "Weight    : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."; Tab(58); Chr(27) + Chr(71); "Bundle(s)   : "; Chr(27) + Chr(72); Trim(rs1!bundles)
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
    kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " order by printorder,sno ", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
                'Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(Str(kk!quantity)), 5); Tab(T6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                Print #1, Tab(0); rsets(Trim(str(sno)), 4); Tab(6); Trim(tdata!Bookname); Tab(41); rsets(Trim(str(kk!quantity)), 5); Tab(48); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(56); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CashB where invoiceno=" + Trim(rs1!INVOICENO) + " and printorder =" + Trim(str(cdiscount)) + " group by printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(56); rsets(Trim(Format(str(tdata(0)), "0.00")), 12)
                   'Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(Str(tdata(0) * cdiscount / 100), "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata(0) - (tdata(0) * cdiscount / 100)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(str(vdis), "0.00")) + " %"; Tab(56); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(69); rsets(Trim(Format(str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(57); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.close
             Loop
         End If
    End If
    Print #1, repli("-", 81)
    Print #1, Tab(39); rsets(Trim(str(totalquantity)), 7); Tab(69); rsets(Trim(Format(str(netamount), "0.00")), 12)
    Line = Line + 2
    If kk.State = 1 Then kk.close
    kk.Open "Select * from CASHC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
       Do While Not kk.EOF
             If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(48); Trim(kk!Text) + "    " + Trim(Format(str(kk!rate), "0.00")); Tab(69); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(48); Trim(kk!Text); Tab(69); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
          
        End If
        Print #1, Tab(69); repli("-", 12)
        Print #1, Chr(27) + Chr(71); Tab(49); "NET AMOUNT  : "; Tab(70); rsets(Trim(Format(str(netamount), "0.00")), 12); Chr(27) + Chr(72)
        VNetamt = netamount
        Line = Line + 2
        kk.close
        kk.Open "Select * from CASHA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(48); kk!txt1 & "    :"; Tab(69); rsets(Trim(Format(str(Abs(kk!txt1a)), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + kk!txt1a
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(48); kk!txt2 & " :"; Tab(69); rsets(Trim(Format(str(Abs(kk!txt2a)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + kk!txt2a
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(48); "CASH RECD.  :"; Tab(69); rsets(Trim(Format(str(Abs(kk!baa)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - kk!baa
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(69); repli("-", 12)
                 Print #1, Tab(48); Chr(27) + Chr(71); "BALANCE     : "; Tab(70); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
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
        tempdata.Open "setup1", CON, adOpenDynamic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(50); "FOR " + Trim(tempdata!CNAME)
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
                    ctl.Text = ""
                End If
                ctl.Enabled = False
            End If
        Next
        For I = 1 To maxrow
           Grid1.row = I
            For J = 1 To 8
                Grid1.col = J
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
        Unload OTHERCASH
End Sub
Public Function templost() As Boolean
    Dim check As Boolean
    Dim row, col As Integer
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
    RRR = Grid1.row
    CCC = Grid1.col
    Grid1.row = lastrow
    Grid1.col = lastcol
    mprevcol = Grid1.col
    Select Case Grid1.col
            Case 1
                Grid1.Text = tempmeb.Text
                '/*************************
                If RS.State = 1 Then
                    RS.close
                End If
                RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                row = Grid1.row
                col = Grid1.col
                If Trim(Grid1.Text) <> "" Then
                    If Not RS.BOF Then
                        RS.MoveFirst
                        RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF Then
                            tempmeb.Visible = True
                            tempmeb.SetFocus
                            RS.close
                            templost = False
                            Exit Function
                        Else
                            Grid1.Text = RS(0)
                            Grid1.col = 2
                            Grid1.Text = RS(1)
                         '   If Not edit Then
                                Grid1.col = 3
                                If Trim(Grid1.Text) = "" Then
                                    Grid1.Text = 0
                                End If
                                q = Val(Grid1.Text)
                                Grid1.col = 5
                                If Trim(Grid1.Text) = "" Then
                                Grid1.Text = Format(RS(3), "0.00")            'rs(3)
                                r = RS(3)
                                End If
                                '/******************
                                
                            '' Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
                                
'------------------------------------------
                            category = returnCategory(Trim(RS(2)))
                            If Optioncash.value = True Then
                            
                            If category = "C1" Then
                               Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                            ElseIf category = "C2" Then
                            'Set kk = con.Execute("select Category2 from sledger where subledger='" + Trim(customercode.Text) + "'")
                               Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                            ElseIf category = "C3" Then
                            'Set kk = con.Execute("select Category2 from sledger where subledger='" + Trim(customercode.Text) + "'")
                               Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII1.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                            End If
                            
                            Else
                            
                            If category = "C1" Then
                               Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
                            ElseIf category = "C2" Then
                               Set kk = CON.Execute("select CATEGORY2 from sledger where subledger='" + Trim(customercode.Text) + "'")
                            ElseIf category = "C3" Then
                               Set kk = CON.Execute("select CATEGORY3 from sledger where subledger='" + Trim(customercode.Text) + "'")
                            End If
                                
                                
                            End If
'-----------------------------------
                                
                                
                                Grid1.col = 6
                                
                                 If kk.BOF Then
                                             GoTo abc
                                 End If
                                       
                                
                                If Grid1.Text = "" And addmode = True Then
                                    If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        
                                kk.close
                                If category = "C1" Then
                                    If Optioncash.value = True Then
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                ElseIf category = "C2" Then
                                    If Optioncash.value = True Then
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                ElseIf category = "C3" Then
                                    If Optioncash.value = True Then
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII1.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                End If
                                        
                                        
                                        Grid1.col = 4
                                        If kk.BOF Then
                                             GoTo abc
                                        End If
                                        Grid1.Text = Format(kk(0), "0.00")
                                        Grid1.col = 6
                                        Grid1.Text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = RS(3)
                                     Else
abc:
                                        Grid1.col = 4
                                        Grid1.Text = Format(RS(4), "0.00")
                                        Grid1.col = 6
                                        Grid1.Text = Format(RS(4), "0.00")
                                        D = RS(4)
                                    End If
                                    
                                    Grid1.col = 7
                                    Grid1.Text = Format(Round(q * r, 2), "0.00")
                                    Grid1.col = 8
                                    Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
                            Else
                                                    
                            If Grid1.Text = "" And addmode = False Then
                                If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    
''                                    'kk.Close
''                                    'If Optioncash.Value = True Then
''                                    '   Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
''
''                                    'Else
''                                    '    Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
''                                    'End If
''
''
''                                    If category = "Category" Then
''                                        kk.Close
''                                        If Optioncash.Value = True Then
''                                            Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
''                                        Else
''                                             Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(kk(0)) + "' and groupcode='" + Trim(RS(2)) + "'")
''                                       End If
''                                    End If

                                    
                                kk.close
                                If category = "C1" Then
                                    
                                    If Optioncash.value = True Then
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                
                                ElseIf category = "C2" Then
                                    
                                    If Optioncash.value = True Then
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                
                                ElseIf category = "C3" Then
                                    
                                    If Optioncash.value = True Then
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII1.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                        
                                        
                                 End If
                                    
                                    
                                    
                                    
                                    
                                    Grid1.col = 4
                                    If kk.BOF Then
                                        GoTo abc
                                    End If
                                    Grid1.Text = Format(kk(0), "0.00")
                                    Grid1.col = 6
                                    Grid1.Text = Format(kk(0), "0.00")
                                    D = kk(0)
                                    r = RS(3)
                            
                                End If
                            End If
                            End If
                            Grid1.col = col
                            RS.close
                        End If
                    End If
                End If
            Case 3, 5, 6
                If Grid1.col <> 3 Then
                    Grid1.Text = Format(Trim(tempmeb.Text), "0.00")
                Else
                    Grid1.Text = Format(Trim(tempmeb.Text), "0")
                End If
                If Trim(Grid1.Text) = "" Then
                    Grid1.Text = 0
                End If
                row = Grid1.row
                col = Grid1.col
                Grid1.col = 3
                q = Val(Trim(Grid1.Text))
                Grid1.col = 5
                r = Val(Trim(Grid1.Text))
                Grid1.col = 6
                D = Val(Trim(Grid1.Text))
                Grid1.col = 7
                Grid1.Text = Format(Round(q * r, 2), "0.00")
                Grid1.col = 8
                Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
                Grid1.col = col
            Case 4
                Grid1.Text = tempmeb.Text
                If Trim(Grid1.Text) = "" Then
                    Grid1.Text = 0
                End If
        End Select
        row = Grid1.row
        col = Grid1.col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            Grid1.row = I
            Grid1.col = 7
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
        Next
        invoicecalc
        Me.tqu.Caption = ""
        For I = 1 To maxrow
            Grid1.col = 3
            Grid1.row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
        Next
        Grid1.row = RRR
        Grid1.col = CCC
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
        mprevcol = Grid1.col
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        Select Case Grid1.col
            Case 2
                Dim row, col As Integer
                row = Grid1.row
                col = Grid1.col
                If Trim(Bookname.Text) = "" Then
                    Grid1.col = 1
                    If Trim(Grid1.Text) = "" Then
                        Grid1.Text = Bookname.Text
                           Bookname.SetFocus
  '********* vk
                          
                          
                          If Trim(Grid1.Text) = "" And row = 1 Then
                                 Grid1.col = 2
                                 Grid1.Text = ""
                                 If Trim(Grid1.Text) = "" Then
                                           
                                          Grid1.col = 1
                                          Bookname.SetFocus
                                          Grid1.SetFocus
                                       Exit Sub
                                 End If
                           End If
              '********
                           Commandother.SetFocus
                           'station.SetFocus
                           
                        Exit Sub
                    End If
                End If
                Grid1.row = row
                Grid1.col = col
                Grid1.Text = Bookname.Text
                '/*************************
                If RS.State = 1 Then
                    RS.close
                End If
                RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                row = Grid1.row
                col = Grid1.col
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
                            
                            Grid1.col = 1
                            Grid1.Text = RS(0)
                            Grid1.col = 2
                            Grid1.Text = RS(1)
                        '    If Not edit Then
                                 Grid1.col = 3
                                If Trim(Grid1.Text) = "" Then
                                        Grid1.Text = 0
                                End If
                                q = Val(Grid1.Text)
                                Grid1.col = 5
                                Grid1.Text = Format(RS(3), "0.00")
                                r = RS(3)
                                '/******************
                                
                            '' Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
                                
''                            category = returnCategory(Trim(RS(2)))
''                            If category = "Category" Then
''                            Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
''                            Else
''                            Set kk = con.Execute("select Category2 from sledger where subledger='" + Trim(customercode.Text) + "'")
''                            End If
 
'------------------------------------------
                            category = returnCategory(Trim(RS(2)))
                            If Optioncash.value = True Then
                            
                                If category = "C1" Then
                                   Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                ElseIf category = "C2" Then
                                   Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                ElseIf category = "C3" Then
                                   Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII1.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                End If
                                
                            
                            Else
                            
                            If category = "C1" Then
                               Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
                            ElseIf category = "C2" Then
                               Set kk = CON.Execute("select CATEGORY2 from sledger where subledger='" + Trim(customercode.Text) + "'")
                            ElseIf category = "C3" Then
                               Set kk = CON.Execute("select CATEGORY3 from sledger where subledger='" + Trim(customercode.Text) + "'")
                            End If
                                
                                
                            End If
'-----------------------------------
 
                                   If kk.BOF Then
                                      GoTo abc
                                   End If
 
 
                                
                                Grid1.col = 6
                                If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    kk.close
                                If category = "C1" Then
                                    If Optioncash.value = True Then
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                
                                ElseIf category = "C2" Then
                                
                                    If Optioncash.value = True Then
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                   
                                    End If
                                        
                                ElseIf category = "C3" Then
                                
                                    If Optioncash.value = True Then
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cboCatII1.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                        
                                        
                                End If
                                
                                    
                                    
                                    
                                 ' Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
                                   Grid1.col = 4
                                   If kk.BOF Then
                                      GoTo abc
                                   End If
                                    Grid1.Text = Format(kk(0), "0.00")
                                    Grid1.col = 6
                                    Grid1.Text = Format(kk(0), "0.00")
                                    D = kk(0)
                                Else
abc:
                                    Grid1.col = 4
                                    Grid1.Text = Format(RS(4), "0.00")
                                    Grid1.col = 6
                                    Grid1.Text = Format(RS(4), "0.00")
                                    D = RS(4)
                                End If
                                Grid1.col = 7
                                Grid1.Text = Round(q * r, 2)
                                Grid1.col = 8
                                Grid1.Text = Round((q * r) * (D / 100), 2)
                         '   End If
                            Grid1.col = col
                            RS.close
                        End If
                    End If
                End If
        End Select
        row = Grid1.row
        col = Grid1.col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            Grid1.row = I
            Grid1.col = 7
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
        Next
        invoicecalc
        Grid1.row = row
        Grid1.col = col
        Select Case Grid1.col
            Case 1
                Grid1.col = 3
                Grid1.SetFocus
                Grid1_Click
            Case 2
                Grid1.col = 3
                Grid1.SetFocus
                Grid1_Click
            Case 3, 4, 5
                Grid1.col = Grid1.col + 1
                Grid1.SetFocus
                Grid1_Click
            Case 6
                Grid1.col = 1
                Grid1.row = Grid1.row + 1
                Grid1.SetFocus
                Grid1_Click
        End Select
    End If
End Sub
Private Sub Bookname_LostFocus()
    Bookname.Visible = False
End Sub

Private Sub bundles_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If Trim(Me.customercode.Text) <> "" Then
            Me.Grid1.col = 1
            Me.Grid1.row = 1
            Me.Grid1.SetFocus
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
If cmbAgentName.Text = "" Then
   MsgBox "Enter a Agent Name.. "
   cmbAgentName.SetFocus
   Exit Sub
Else
  Dim rs1 As New ADODB.Recordset
  rs1.Open "select *  from AgentMaster where AgentName='" & cmbAgentName.Text & "' order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
  If rs1.RecordCount <= 0 Then
     MsgBox "Enter Valid Agent Name.. "
     cmbAgentName.SetFocus
  End If
End If

End Sub

Private Sub cmbareaname_LostFocus()
  Me.textbox.Text = Me.textbox.Text + ", " + cmbareaname.Text
  cmbareaname.Visible = False
End Sub

Private Sub Combosldistrictcode_LostFocus()
If Combosldistrictcode.Text = "" Then
   Combosldistrictcode.SetFocus
   Exit Sub
End If
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

If Combosldistrictcode.Text <> "" Then
   rs1.Open "Select * from Districts where Districtname = '" & Combosldistrictcode.Text & "'", CON, adOpenStatic, adLockReadOnly
   If rs1.RecordCount <= 0 Then
      MsgBox "Please Select valid district.."
      Combosldistrictcode.SetFocus
   End If
End If
Set rs1 = New ADODB.Recordset
If Combosldistrictcode.Text <> "" And addmode = True Then
   rs1.Open "Select * from Districts where Districtname = '" & Combosldistrictcode.Text & "'", CON, adOpenStatic, adLockReadOnly
   If rs1.RecordCount > 0 Then
      Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
   End If
End If





End Sub

Private Sub Command1_Click()
Printheader = True
   printinvoice
End Sub

Private Sub Commandprintnh_Click()

    printch = "CASHA"
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
'On Error Resume Next
    invoiceabandon
    Dim RS As ADODB.Recordset
    addoredit = True
    addmode = True
    edit = False
    Set RS = New ADODB.Recordset
    Dim TEMPNUM As Integer
    
    If edit = False Then
    'If CON.Execute("Select max(invoiceno) from CASHA")(0) >= Val(Trim(Me.I_NO.Text)) Then
         RS.Open "Select max(invoiceno) from CASHA", CON, adOpenDynamic, adLockOptimistic
         If IsNull(RS(0)) Then
           Me.I_NO.Text = 1
         Else
           Me.I_NO.Text = RS(0) + 1
         End If
         
         
         RS.Update
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
    
    Grid1.Enabled = True
    Me.customercode.Enabled = True
    Me.Optioncash.SetFocus
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
    Grid1.col = 1
    Grid1.row = 1
    If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select * from books order by BOOKCODE", CON, adOpenDynamic, adLockReadOnly, adCmdText
    row = Grid1.row
    col = Grid1.col
    If Not RS.BOF Then
        RS.MoveFirst
        Do While Not RS.EOF
            Grid1.col = 1
            Grid1.Text = RS(0)
            Grid1.col = 2
            Grid1.Text = RS(1)
            Grid1.col = 3
            If Trim(Grid1.Text) = "" Then
                Grid1.Text = Val(myvalue)
            End If
            q = Val(Grid1.Text)
            Grid1.col = 5
            Grid1.Text = Format(RS(3), "0.00")            'rs(3)
            r = RS(3)
            '/******************
            
            
           category = returnCategory(Trim(RS(2)))
           If category = "C1" Then
            Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
           ElseIf category = "C2" Then
            Set kk = CON.Execute("select Category2 from sledger where subledger='" + Trim(customercode.Text) + "'")
           ElseIf category = "C3" Then
            Set kk = CON.Execute("select Category3 from sledger where subledger='" + Trim(customercode.Text) + "'")
           End If

            
            ''Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
            
            Grid1.col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.close
                Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
                Grid1.col = 4
                If kk.BOF Then
                    GoTo abc
                End If
                Grid1.Text = Format(kk(0), "0.00")
                Grid1.col = 6
                Grid1.Text = Format(kk(0), "0.00")
                D = kk(0)
            Else
abc:
                Grid1.col = 4
                Grid1.Text = Format(RS(4), "0.00")
                Grid1.col = 6
                Grid1.Text = Format(RS(4), "0.00")
                D = RS(4)
            End If
            Grid1.col = 7
            Grid1.Text = Format(Round(q * r, 2), "0.00")
            Grid1.col = 8
            Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
            If Not RS.EOF Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.row = Grid1.row + 1
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
    For I = 1 To Grid1.Rows - 1
            Grid1.row = I
            Grid1.col = 7
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
            Grid1.col = 3
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
     Next
     maxrow = Grid1.Rows - 1
Else
'Grid1_Click
Exit Sub
End If

invoicecalc
txtMark.ListIndex = 0

End Sub

Private Sub Commanddelete_Click()


    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
   
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from casha where invoiceno=" & I_NO.Text & "", CON
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select * from casha where invoiceno=" & I_NO.Text & "", CON
       'If rs_h.Fields("Print_yes").Value = "y" Then
          If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
          End If
       'End If
       
    End If
  



If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
                CON.Execute ("delete  from CASHA where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete  from CASHB where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete  from CASHC where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                invoiceabandon
End If
End Sub

Private Sub Commandedit_Click()
   If I_NO.Text <> "" Then
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
    Grid1.Enabled = True
    Commandall.Enabled = False
    Me.customercode.Enabled = True
    edit = True
    I_NO_LostFocus
    i_dt.Enabled = True
    i_dt.SetFocus
    ' CASHCTMP creation start
    DoEvents
    CON.Execute ("Delete  from CASHCTMP")
    DoEvents
    CON.Execute ("insert into CASHCTMP(INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate,AMOUNT,DEBITORCREDIT,[TEXT],RYN,Fyear,setupid)  select INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate,AMOUNT,DEBITORCREDIT,[TEXT],RYN,Fyear,setupid from CASHC where INVOICENO = " + Trim(I_NO.Text))
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
    OTHERCASH.Top = 0
    OTHERCASH.Left = 0
    OTHERCASH.Visible = False
    
End Sub
Private Sub Commandother_Click()

Commandsave.Enabled = True
searchForm = "cash"
frmEndPartTrans.Show
    
End Sub
Private Sub CommandPrint_Click()
   
   
printch = "casha"
ino = I_NO
printch1 = "INVOICENO"
Printheader = True
printinvoice

End Sub
Private Sub Commandreturn_Click()
''   Dim rs As New ADODB.Recordset
''   rs.Open "tempCASH", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
''   If rs.BOF Then
''       rs.AddNew
''   End If
''   rs!In = CON.Execute("Select max(invoiceno) from cashA")(0)
''   rs.Update
''   rs.Close

Unload Me
addoredit = False

'MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()
     
    
    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
   
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from casha where invoiceno=" & I_NO.Text & "", CON
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select * from casha where invoiceno=" & I_NO.Text & "", CON
       'If rs_h.Fields("Print_yes").Value = "y" Then
          If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
          End If
       'End If
       
    End If
  
  
  
  Dim RS As ADODB.Recordset
  Set RS = New ADODB.Recordset

  
  
  If edit = False And addmode = False Then
      Me.Commandsave.Enabled = False
      Exit Sub
    End If


If Optioncash = True Then
    If Trim(Combosldistrictcode) = "" Then
        MsgBox "Please Enter District"
        Exit Sub
    End If
    If Val(Trim(Me.mna.Caption)) <> Val(Trim(frmEndPartTrans.T3TEXT.Text)) Then
      MsgBox "In This Bill  Netamount and Cash Reciept Are Not Equal." + Chr(13) + "Please Select Ctedit Option For Part Cash Memo"
      Exit Sub
    End If
   

    
    
Else
   If Trim(cmbAgentName.Text) = "" Then
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
Grid1.row = 1
Grid1.col = 1
If Trim(Grid1.Text) = "" Then
   MsgBox "Please Enter item.... "
   Exit Sub
End If


'----------------------------------------------------------------
'I_NO.Text = 852
If edit = False Then
   If check_Duplikate("casha", I_NO.Text) = True Then
      MsgBox "This  Inv. Number Already Exist ..", vbCritical
      Exit Sub
   End If
End If
'----------------------------------------------------------------


If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(customercode.Text) <> "" Then
   
   If edit Then
      CON.Execute ("delete  from CASHA where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
      CON.Execute ("delete  from CASHB where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
      CON.Execute ("delete  from CASHC where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
      CON.Execute ("delete  from CashRegister where cmno = " + Trim(I_NO.Text))
   End If
   
   If RS.State = 1 Then RS.close
   LAMOUNT = 0
   
 If edit = False Then
 If (I_OB <> "" And txtMark <> "") Then
    Party_Remove_FromOrder Trim(Me.customercode.Text), txtMark, Trim(I_OB)
 End If
 End If
 
   RS.Open "select * from CASHA where  invoiceno<=0", CON, adOpenDynamic, adLockOptimistic
   If Not edit Then
again:
      If CON.Execute("Select max(invoiceno) from CASHA")(0) >= Val(Trim(Me.I_NO.Text)) Then
      
      End If
   End If
   RS.AddNew
   RS!INVOICENO = Val(Me.I_NO.Text)
   RS!invoicedate = Me.i_dt.Text
   RS!Genledger = Trim(Me.Genledger.Text)
   RS!SUBLEDGER = Trim(Me.customercode.Text)
   RS!ORDERBY = Trim(Me.I_OB.Text)
   If Trim(Me.I_DTOB) = Trim("__/__/____") Then
     
      RS!ORDERDATE = Null
   Else
      RS!ORDERDATE = Trim(Me.I_DTOB.Text)
   End If
   'rs!marka = Trim(Me.marka.Text)
   RS!bundles = Trim(Me.bundles)
   RS!Godown = txtMark.Text
   'rs!through1 = Trim(Me.through1.Text)
   'If Trim(Me.through1.Text) = "" Then
   '   rs!through1 = " "
   'End If
   RS!station = Trim(Me.station.Text)
   RS!biltyno = Trim(Me.biltno.Text)
   If Trim(Me.bdated) = Trim("__/__/____") Then
      RS!BILTYDATE = Null
      'rs!BILTYDATE = Date
   Else
     RS!BILTYDATE = Me.bdated & ""
   End If
   RS!transportname = Trim(Me.cmbtransportname.Text)
   RS!freight = Me.freight & ""
  ' rs!weight = Me.weight & ""
   RS!netamount = Round(Val(Trim(Me.mna.Caption)), 2)
   RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
   RS!txt1 = Trim(frmEndPartTrans.T1TEXT.Text)
   RS!txt1a = Val(Trim(frmEndPartTrans.T1.Text))
   RS!txt2 = Trim(frmEndPartTrans.T2TEXT.Text)
   RS!txt2a = Val(Trim(frmEndPartTrans.T2.Text))
   'rs!baa = Val(Trim(frmEndPartTrans.T3TEXT.Text))
   
   RS!baa = Val(Trim(frmEndPartTrans.T3TEXT.Text))
   RS!baa = Val(Trim(labelbybank.Caption))
   
   RS!District = Combosldistrictcode.Text
   RS!CASHPARTYNAME = textbox.Text
   RS!agentname = cmbAgentName.Text
   RS!discat = cmbdiscountcat.Text
   RS!discatII = cboCatII.Text
   RS!discatIII = cboCatII1.Text
   
err1:
   If Not edit Then
      If CON.Execute("Select max(invoiceno) from CASHA")(0) >= Val(Trim(Me.I_NO.Text)) Then
         'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
         'rs!INVOICENO = Val(Me.I_NO.Text)
         On Error GoTo err1
      End If
   End If
               
   RS!fyear = session
   RS!setupid = setupid

   RS.Update
   On Error GoTo 0
   RS.close
   RS.Open "select * from CASHB where invoiceno<=0", CON, adOpenDynamic, adLockOptimistic
   Dim I As Integer
   RRRR = Grid1.row
   CCCC = Grid1.col
   For I = 1 To maxrow
       Grid1.row = I
       Grid1.col = 1
       If Trim(Grid1.Text) <> "" Then
          Grid1.col = 3
          If Val(Trim(Grid1.Text)) > 0 Then
             Grid1.col = 5
            If Val(Trim(Grid1.Text)) > 0 Then
               RS.AddNew
               Grid1.col = 1
               RS!INVOICENO = Val(Me.I_NO.Text)
               RS!invoicedate = Me.i_dt.Text
               RS!Genledger = Trim(Me.Genledger.Text)
               RS!SUBLEDGER = Trim(Me.customercode.Text)
               RS!Bookcode = Trim(Grid1.Text)
               Grid1.col = 3
               RS!quantity = Trim(Grid1.Text)
               Grid1.col = 5
               RS!rate = Trim(Grid1.Text)
               Grid1.col = 7
               RS!amount = Trim(Grid1.Text)
               LAMOUNT = Val(Trim(Grid1.Text))
               Grid1.col = 4
               RS!PRINTORDER = Trim(Grid1.Text)
               Grid1.col = 6
               RS!discount = Trim(Grid1.Text)
               Grid1.col = 8
               RS!netamount = LAMOUNT - Trim(Grid1.Text)
               LAMOUNT = 0
               RS!agentname = Trim(Me.cmbAgentName.Text)
               RS!fyear = session
               RS!setupid = setupid

               RS.Update
            End If
         End If
     End If
  Next
  RS.close
  Grid1.TopRow = 1
  RS.Open "select * from CASHC where invoiceno<=0 ", CON, adOpenDynamic, adLockOptimistic
  '/******
  'Dim I, x As Integer
  
  
  
   Dim temprs As ADODB.Recordset
   Set temprs = New ADODB.Recordset
       For I = 1 To frmEndPartTrans.vs.Rows - 1
           frmEndPartTrans.vs.row = I
           frmEndPartTrans.vs.col = 0
           If Trim(frmEndPartTrans.vs.Text) <> "" Then
              RS.AddNew
              RS!INVOICENO = Val(Me.I_NO.Text)
              RS!invoicedate = Me.i_dt.Text
              RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
              RS!Text = Trim(frmEndPartTrans.vs.Text)
              If temprs.State = 1 Then temprs.close
              If edit Then
                 temprs.Open "select * from CASHCTMP WHERE INVOICENO = " & Val(Me.I_NO.Text) & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
                 If frmEndPartTrans.vs.Text <> "" Then
                    temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.Text) + "'"
                    RS!Genledger = temprs!Genledger & ""
                    RS!SUBLEDGER = temprs!SUBLEDGER & ""
                    RS!DebitorCredit = temprs!DebitorCredit & ""
                    RS!RYN = temprs!RYN & ""
                End If
                temprs.close
              Else
                 temprs.Open "select * from INVOICEEND where type='" & searchForm & "' and  " & stringyear & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
                 If frmEndPartTrans.vs.Text <> "" Then
                    temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.Text) + "'"
                    RS!Genledger = temprs!Genledger & ""
                    RS!SUBLEDGER = temprs!SUBLEDGER & ""
                    RS!DebitorCredit = temprs!DebitorCredit & ""
                    RS!RYN = temprs!RYN & ""
                 End If
                 temprs.close
              End If
              frmEndPartTrans.vs.col = 1
              RS!rate = Val(Trim(frmEndPartTrans.vs.Text))
              If Val(Trim(frmEndPartTrans.vs.Text)) > 0 Then
                 RS!amount = Round((Me.totalamount - Me.totaldiscount), 2) * Round((Val(Trim(frmEndPartTrans.vs.Text)) / 100), 2)
              Else
                frmEndPartTrans.vs.col = 2
                RS!amount = Val(Trim(frmEndPartTrans.vs.Text))
              End If
              RS!fyear = session
              RS!setupid = setupid

              RS.Update
          End If
      Next
      RS.close
      
      CON.Execute ("delete  from CASHCTmp where INVOICENO = " + Trim(I_NO.Text))
      
  
  SAVED = True
  
  End If
  
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

    
     UpdateDisPatchReg1 I_NO, i_dt, Me.customercode, PopUpValue1, Trim(Me.bundles), Trim(Me.cmbtransportname.Text), "-", Trim(Me.biltno.Text), Me.bdated, Trim(Me.freight), "CashRegister"
     PopUpValue1 = ""
    End If
    'End If

  
  If SAVED Then
      Unload frmEndPartTrans
   
      MsgBox "Record Saved"
      
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
 End Sub

Private Sub Commandsearch_Click()
   '''Unload frmEndPartTrans
   '''Me.Enabled = False
   '''searchscreen.Grid1.row = 0
   '''searchscreen.Grid1.col = 0
   searchType = "inv"
   popuplist10 "select InvoiceNo,InvoiceDate,Subledger,NetAmount from CASHA where " & stringyear & "  order by InvoiceNo", CON

End Sub

Private Sub Commandsearch_GotFocus()
If PopUpValue1 <> "" Then
     I_NO.Text = PopUpValue1
     I_NO_LostFocus
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
        RS.Open "select * from sledger where gledger='SUNDRY DEBTORS' and subledger='" + Trim(customercode.Text) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.RecordCount <= 0 Then
           customercode.SetFocus
           HIT
           RS.close
           Exit Sub
        End If
        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        If RS!distcode <> "" And addmode = True Then
            rs1.Open "Select * from Districts where Districtname = '" & RS!distcode & "'", CON, adOpenStatic, adLockReadOnly
            If rs1.RecordCount > 0 Then
                Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
            End If
             Combosldistrictcode.Text = RS!distcode
        Else
        Combosldistrictcode.Text = RS!distcode
        End If
        Me.textbox.Text = Me.customercode.Text
        Me.customercode.Visible = False

End Sub

Private Sub Delete_Click()
If Grid1.row >= 1 Then
    Grid1.SetFocus
    Grid1.RemoveItem (Grid1.row)
    If Grid1.row > 1 Then
        Grid1.row = Grid1.row - 1
    End If
    Grid1_Click
End If
End Sub

Private Sub Form_Activate()
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    mna.Enabled = True
    Label2.Enabled = True
    'txtMark.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If Grid1.row >= 1 Then
           Grid1.RemoveItem Grid1.row
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
          If addmode = True Then
                SendKeys "{DOWN}"
           End If
            SendKeys "{TAB}"
        Else
            If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("grid1")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("tempmeb")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bookname")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bundles")) Then
                SendKeys ("{TAB}")
            End If
        End If
    End If
End Sub


Private Sub Form_Load()
On Error Resume Next

      
    BackColorFrom Me
    addmode = False
    edit = False
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
    Me.Top = 0
    Me.Left = 0
      
    Grid1.Rows = 2
    Grid1.Cols = 1
    Grid1.Rows = 10
    Grid1.Cols = 9
    Grid1.row = 0
    Grid1.col = 1
    Grid1.Text = "Book Code "
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Book Name"
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Quantity"
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Print. Ord."
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Rate"
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Disc %"
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Amount"
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Disc. Amount"
    Grid1.RowHeight(0) = Grid1.CellHeight + 50
    Grid1.ColWidth(0) = 150
    Grid1.ColWidth(1) = 1100
    Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 750
    Grid1.ColWidth(4) = 750
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 1200
    Grid1.ColWidth(8) = 1200
    Me.CommandPrint.Enabled = True
    Me.Commandprintnh.Enabled = True
    RS.Open "select * from books", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
    RS.Open "select Distinct categorycode from DISCCATS order by categorycode", CON, adOpenDynamic, adLockReadOnly, adCmdText
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

    
    
    Genledger.Text = "SUNDRY DEBTORS"
    RS.Open "select * from sledger where gledger='" + Trim(Genledger.Text) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            Me.customercode.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    
     '*******Agent  combo fill
    RS.Open "select  Agentname from AgentMaster order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
 
    
    
      Dim rs_godwn As New ADODB.Recordset
    
    If rs_godwn.State = 1 Then rs_godwn.close
    rs_godwn.Open "select godwn from GodownMaster order by id", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
    
    'RS.Open "tempcash", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
    RS.Open "SELECT MAX(INVOICENO) FROM CASHA", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not IsNull(RS(0)) Then
       Me.I_NO.Text = RS(0) + 1
       countersale.Enabled = True
       countersale.edit = False
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
       SendKeys "{TAB}"
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
    Picture5.Enabled = True
    If RS.State = 1 Then RS.close
    RS.Open "select * from DISTRICTS order by DISTRICTNAME", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS!DISTRICTNAME
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If

    
    If RS.State = 1 Then RS.close
    '''RS.Open "select * from Area order by areaNAME", CON, adOpenDynamic, adLockReadOnly, adCmdText
    RS.Open "select * from DISTRICTS order by DISTRICTNAME", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.cmbareaname.AddItem RS!DISTRICTNAME
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub

Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub
Function returnCategory(s As String) As String
    Dim s1 As New ADODB.Recordset
    If s1.State = 1 Then s1.close
    
    s1.Open "select category from [groups] where groupcode='" & s & "'", CON
    If s1.EOF = False Then
       returnCategory = s1(0)
    End If
    
End Function


Sub Grid1_Click()
If Trim(Me.customercode.Text) <> "" Then
Dim PREVROW As Integer
Dim prevcol As Integer
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
prevcol = Me.Grid1.col
PREVROW = Me.Grid1.row

If Me.Grid1.row > 1 Then
    Grid1.row = Grid1.row - 1
    Grid1.col = 1
    If Trim(Grid1.Text) <> "" Then
        Grid1.row = PREVROW
        Grid1.col = prevcol
        If Trim(Me.customercode.Text) <> "" Then
            If Me.customercode.Enabled = True Then
                Me.customercode.Enabled = False
            End If
            Grid1.col = 1
            If prevcol > 1 And Trim(Grid1.Text) = "" Then
                Grid1.col = 2
                SendKeys Chr(13)
            Else
                Grid1.col = prevcol
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
        Me.Grid1.col = 1
        If prevcol > 1 And Trim(Grid1.Text) = "" Then
            Me.Grid1.col = 2
            Me.Grid1.SetFocus
            SendKeys Chr(13)
        Else
        'IF GRID1.COL
            Me.Grid1.col = prevcol
            Me.Grid1.SetFocus
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
            
            Select Case Grid1.col
            Case 1, 3, 4, 5, 6
                Bookname.Visible = False
                tempmeb.Visible = True: tempmeb.Enabled = True
                tempmeb.ZOrder
                If Grid1.col <> 1 Then
                    If Grid1.col <> 3 Then
                        tempmeb.Text = Format(Grid1.Text, "0.00")
                        
                    Else
                        tempmeb.Text = Format(Grid1.Text, "0")
                    End If
                   
                Else
                    tempmeb.Text = Grid1.Text
                End If
                tempmeb.Width = Grid1.ColWidth(Grid1.col)
                tempmeb.Left = Grid1.CellLeft + 80
                tempmeb.Top = Grid1.Top + Grid1.CellTop '- 50
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.Text = Grid1.Text
                Bookname.Top = Grid1.Top + Grid1.CellTop
                Bookname.Left = Grid1.CellLeft + 80
                Bookname.Width = Grid1.ColWidth(Grid1.col)
            Case 6
                
            Case Else
                Bookname.Visible = False
                tempmeb.Visible = False
            End Select
            Select Case Grid1.col
                Case 1, 3, 4, 5, 6
                    tempmeb.Mask = ""
                    tempmeb.MaxLength = 20
                Case 2
                    With Bookname
                        .Visible = True
                        .ZOrder
                    End With
                End Select
            Select Case Grid1.col
            Case 2
                Bookname.SetFocus
                If KeyAscii <> 13 Then
                    SendKeys Chr(KeyAscii)
                End If
            Case 1, 3, 4, 5, 6
                mprevcol = Grid1.col
                tempmeb.SetFocus
            Case Else
                If KeyAscii = 13 Then
                    SendKeys "{RIGHT}"
                End If
            End Select
        End If
    If maxrow < Grid1.row Then
        maxrow = Grid1.row
    End If
End If
    lastrow = Grid1.row
    lastcol = Grid1.col
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
        Grid1.SetFocus
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
    trs2.Open "Select invoiceno as cn from casha", CON, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount <= 0 Then
       Exit Sub
    Else
        If tRS1.State = 1 Then tRS1.close
        tRS1.Open "Select min(invoiceno) as mid,invoicedate from casha group by invoiceno,invoiceDate", CON, adOpenDynamic, adLockOptimistic
        If tRS1.RecordCount > 0 Then
             
            If CDate(i_dt) <= tRS1!invoicedate Then
            
               If CDate(i_dt) <> tRS1!invoicedate Then
               If Month(CDate(i_dt)) <> 4 And Day(CDate(i_dt)) <> 1 Then
                 MsgBox "Please select valid Cash Memo No. for this date.."
                 i_dt.SetFocus
                 Exit Sub
                   
            Else
                If tRS1!Mid <> 1 Then
               If Val(I_NO) >= tRS1!Mid Then
                 MsgBox "Please select Cash Memo No. for this date.."
                 i_dt.SetFocus
                 Exit Sub
               End If
               End If
                 End If
                 
               End If
            End If
        End If
    End If
    
    
    
    If trs2.State = 1 Then trs2.close
    trs2.Open "Select max(invoiceno) as mid from casha where  invoicedate <= cdate('" & i_dt.Text & "')-1", CON, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount > 0 Then
        If IsNull(trs2!Mid) <> True Then
            If Val(I_NO.Text) >= trs2!Mid Then
               If tRS1.State = 1 Then tRS1.close
               tRS1.Open "Select  min(InvoiceNo)as m2 from casha where invoicedate >= cdate('" & i_dt.Text & "')+1", CON, adOpenDynamic, adLockOptimistic
               If tRS1.RecordCount > 0 Then
                  If IsNull(tRS1!m2) <> True Then
                     If Val(I_NO.Text) <= tRS1!m2 Then
                       
                     Else
                         MsgBox "Please select valid Cash Memo No. for this date.."
                         I_NO.SetFocus
                     End If
                  End If
               End If
            
            Else
               If I_NO.Enabled = False Then Exit Sub
                    If i_dt.Enabled = True Then
                            MsgBox "Please select valid Cash Memo No. for this date.."
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

Sub I_NO_LostFocus()
    On Error Resume Next
    
    If Val(inviceNo) > 0 Then
       I_NO.Text = inviceNo
    End If
    
    inviceNo = ""
    
    
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    If Trim(I_NO.Text) = "" Then
        MsgBox "Cash Memo No cannot be null"
        I_NO.SetFocus
    Else
        If RS.State = 1 Then RS.close
        RS.Open "Select * from  CASHA where INVOICENO = " + Trim(I_NO.Text) + "", CON, adOpenStatic, adLockReadOnly
        If RS.EOF Then
            If addoredit = False Then
            '     MsgBox "Cash Memo No not found"
            '     Exit Sub
            End If
            'Exit Sub
        End If
        
        If addoredit Then
            MsgBox "Cash Memo No already exist..."
            'I_NO.SetFocus
            HIT
            Exit Sub
        End If
        Dim ctl As Control
        For Each ctl In Me.Controls
           If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = True
            End If
        Next
        Me.Commandother.Enabled = True
        I_NO.Text = RS!INVOICENO
        Me.i_dt.Text = RS!invoicedate
        Me.Genledger.Text = Trim(RS!Genledger)
        Me.customercode.Text = Trim(RS!SUBLEDGER)
        Me.textbox.Text = Trim(RS!SUBLEDGER)
        Me.I_OB.Text = Trim(RS!ORDERBY)
        Me.I_DTOB.Text = IIf(IsNull(RS!ORDERDATE), "__/__/____", RS!ORDERDATE)
        'Me.marka.Text = Trim(rs!marka)
        Me.bundles = Trim(RS!bundles)
        txtMark = RS!Godown & ""
        'Me.through.Text = rs!through
        'Me.through1.Text = rs!through1
        Me.station.Text = RS!station
        Me.biltno.Text = Trim(RS!biltyno)
        Me.bdated = IIf(IsNull(RS!BILTYDATE), "__/__/____", RS!BILTYDATE)
        Me.freight = Trim(RS!freight)
        'Me.weight = Trim(rs!weight)
        Me.labelbybank = Format(Round(Val(RS!baa), 2), "0.00")
        mna.Caption = Format(Round(Val(RS!netamount), 2), "0.00")
        Me.cmbtransportname.Text = IIf(IsNull(RS!transportname), "", RS!transportname)

      
        
        If RS!District <> "" Then
            Combosldistrictcode.Text = RS!District
        End If
        textbox.Text = RS!CASHPARTYNAME
        If Me.customercode.Text = "CASH PARTY" Then
            Optioncash = True
            Me.cmbAgentName.Text = IIf(IsNull(RS!agentname), "", Trim(RS!agentname))
        Else
            Optioncredit = True
            Me.cmbAgentName.Text = IIf(IsNull(RS!agentname), "", Trim(RS!agentname))
        End If
        cmbdiscountcat.Text = IIf(IsNull(RS!discat), "", RS!discat)
        cboCatII.Text = IIf(IsNull(RS!discatII), "", RS!discatII)
        cboCatII1.Text = IIf(IsNull(RS!discatIII), "", RS!discatIII)
        
        RS.close
        Grid1.TopRow = 1
    '*/**/*/*/*/*//*/*
    If RS.State = 1 Then RS.close
    RS.Open "Select * from CASHB where INVOICENO =" + Trim(I_NO.Text) + "  order by SNO ", CON, adOpenStatic, adLockReadOnly
    If Not RS.EOF Then
            Grid1.row = 1
            Grid1.col = 1
            Do While Not RS.EOF
               If Trim(RS!INVOICENO) = Trim(I_NO.Text) Then
                Grid1.col = 1
                Grid1.Text = Trim(RS!Bookcode)
                If kk.State = 1 Then
                    kk.close
                End If
                kk.Open "select * from books where bookcode='" + Trim(RS!Bookcode) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
                Grid1.col = 2
                Grid1.Text = Trim(kk!Bookname)
                Grid1.col = 3
                Grid1.Text = Trim(RS!quantity)
                Grid1.col = 5
                Grid1.Text = Format(Round(RS!rate, 2), "0.00")
                Grid1.col = 7
                Grid1.Text = Format(Round(RS!amount, 2), "0.00")
                Grid1.col = 4
                Grid1.Text = Format(Round(RS!PRINTORDER, 2), "0.00")
                Grid1.col = 6
                Grid1.Text = Format(Round(RS!discount, 2), "0.00")
                Grid1.col = 8
                Grid1.Text = Format(Round(RS!amount * (RS!discount / 100), 2), "0.00")
                End If
                If Not RS.EOF Then
                    RS.MoveNext
                End If
                Grid1.row = Grid1.row + 1
                Grid1.Rows = Grid1.Rows + 1
            Loop
            maxrow = Grid1.row
        End If
        row = Grid1.row
        col = Grid1.col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            Grid1.row = I
            Grid1.col = 7
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
        Next
        mga.Caption = Format(Round(totalamount, 2), "0.00")
        mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
        Me.tqu.Caption = ""
        For I = 1 To maxrow
            Grid1.col = 3
            Grid1.row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
        Next
        Grid1.row = RRR
        Grid1.col = CCC
    End If
    mna.Enabled = True
    Label2.Enabled = True
    Me.Commandother.Enabled = True
End Sub

Private Sub I_OB_GotFocus()
Dim trs As New ADODB.Recordset
trs.Open " SELECT DISTCODE    FROM SLEDGER  WHERE SUBLEDGER='" & customercode.Text & "'", CON, adOpenStatic, adLockOptimistic, adCmdText
       If Not trs.BOF Then
           If Combosldistrictcode.Text = "" Then
               Combosldistrictcode.Text = IIf(IsNull(trs!distcode), "", trs!distcode)
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
       Label4.Visible = True
       Combosldistrictcode.Visible = True
       'Label4.Left = 4220
       Combosldistrictcode.Enabled = True
       'Combosldistrictcode.Left = 5350
      ' Label11.Left = 6950
      ' bundles.Left = 8040
       cmbdiscountcat.Visible = True
       lbldis(0).Visible = True
       lbldis(1).Visible = True
       cboCatII.Visible = True
       'cmbAgentName.Visible = False
       'Label13.Visible = False
       
       cboCatII1.Visible = True
       lbldis(2).Visible = True

       
 End If


End Sub

Private Sub Optioncredit_Click()
If Optioncredit.value = True Then
       Label4.Visible = False
       Combosldistrictcode.Visible = True
    '   Label11.Left = 4220
      ' bundles.Left = 5350
       lbldis(0).Visible = False
       lbldis(1).Visible = False
       cmbdiscountcat.Visible = False
       'Combosldistrictcode.Left = 8950
       Combosldistrictcode.Visible = False
       cboCatII.Visible = False
       'cmbAgentName.Visible = True
       'Label13.Visible = True
       
       cboCatII1.Visible = False
       lbldis(2).Visible = False


End If

End Sub

Private Sub station_LostFocus()
station = UCase(station)
End Sub

Private Sub tempmeb_Change()
If Grid1.col = 1 Or Grid1.col = 2 Then
    Grid1.Text = tempmeb.Text
Else
    If Grid1.col = 3 Then
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
            Select Case Grid1.col
                Case 1
                    If RS.State = 1 Then
                        RS.close
                    End If
                    RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                    If Not RS.BOF Then
                        RS.MoveFirst
                        RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF And Trim(Grid1.Text) <> "" Then
                            RS.close
                            Exit Sub
                        Else
                            RS.close
                        If Trim(Grid1.Text) <> "" Then
                                Grid1.col = 3
                            Else
                                Grid1.col = 2
                            End If
                        End If
                    Else
                        If Trim(Grid1.Text) <> "" Then
                            Grid1.col = 3
                        Else
                            Grid1.col = 2
                        End If
                    End If
                    Grid1.SetFocus
                    Grid1_Click
                Case 3
                    If Val(tempmeb.Text) > 0 Then
                        Grid1.col = Grid1.col + 2
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                Case 4
                    Grid1.col = Grid1.col + 2
              '       SendKeys "{LEFT}"
                    Grid1.SetFocus
                    Grid1_Click
                Case 5
                    If Val(tempmeb.Text) > 0 Then
                        Grid1.col = Grid1.col - 1
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                Case 6
                    If Val(Grid1.TextMatrix(Grid1.row, 4)) <> Val(Grid1.TextMatrix(Grid1.row, 6)) Then
                      MsgBox "Discount And Printorder  Not Match.."
                      
                   End If
                    Grid1.col = 1
                    Grid1.row = Grid1.row + 1
                    Grid1.Rows = Grid1.Rows + 1
                    Grid1.SetFocus
                    Grid1_Click
            End Select
        Else
        If Grid1.col = 3 Or Grid1.col = 4 Or Grid1.col = 5 Or Grid1.col = 6 Then
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

 If Optioncash = True And textbox.Text <> "" Then
        Me.customercode.Text = "CASH PARTY"
 End If
 textbox.Text = UCase(textbox.Text)
 
End Sub

Private Sub through_LostFocus()
through = UCase(through)
End Sub

Private Sub through1_LostFocus()
through1 = UCase(through1)
End Sub



Private Sub weight_KeyPress(KeyAscii As Integer)
    
 'by vk
 'If KeyAscii = 13 Then
   '     If Trim(Me.customercode.Text) <> "" Then
    '      Grid1.col = 1
    '      Grid1.row = 1
     '     Grid1_Click
     ' Else
     '   Me.textbox.SetFocus
        'Me.customercode.SetFocus
    ' End If
    
    'End If
    
    
    
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
          kkk.Open "select * from setup1", CON, adOpenDynamic, adLockReadOnly, adCmdText
          If Not kkk.BOF Then
             Print #1, Chr(27) + Chr(15) + Chr(14)
             Print #1, Tab(((paperWidth - (Len(Trim(kkk!CNAME)) * 2)) / 2)); Chr(27) + Chr(15) + Chr(14); Trim(kkk!CNAME)
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
    rs1.Open "CASHA", CON, adOpenDynamic, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
   
    Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!SUBLEDGER; Tab(T5); "Cash Memo No. : "; Trim(rs1!INVOICENO); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!invoicedate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.close
            End If
            kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE
                Print #1, Tab(3); kkk!ADDRESS1; Tab(T5); "Order by : "; Trim(rs1!ORDERBY); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!ORDERDATE
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
            kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
                        tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
                        Print #1, rsets(Trim(str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(str(kk!quantity)), 5); Tab(T6 + 1); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
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
                        tdata.Open "select sum(amount) from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenDynamic, adLockReadOnly, adCmdText
                        If Not tdata.BOF Then
                            
                            Print #1, Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0), 2)), "0.00")), 12)
                            Print #1, Tab(T5); "Less Discount @ " + Trim(Format(str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                            netamount = netamount + Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                        End If
                        tdata.close
                        Print #1, Tab(T7); repli("-", 22)
                Loop
            End If
           End If
           Print #1, Tab(T5 - 6); rsets(Trim(str(totalquantity)), 7); Tab(T8 + 5); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12)
           Print #1, Tab(T6); repli("-", 22)
           If kk.State = 1 Then
                kk.close
           End If
           kk.Open "Select * from CASHC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
           If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5); Trim(kk!Text) + "    " + Trim(Format(str(Round(kk!rate, 2)), "0.00")); Tab(T8 + 5); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5); Trim(kk!Text); Tab(T8 + 5); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T6); repli("-", 22)
            Print #1, Tab(6); Chr(71) + "NET AMOUNT: "; Tab(T8 + 5); Chr(72) + rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12)
           End If
           kk.close
           kk.Open "Select * from CASHA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
           If Not kk.BOF Then
                If kk!txt1a <> 0 Then
                    Print #1, Tab(T5); kk!txt1; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                    netamount = netamount + Round(kk!txt1a, 2)
                End If
                If kk!txt2a <> 0 Then
                    Print #1, Tab(T5); kk!txt2; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                    netamount = netamount + Round(kk!txt2a, 2)
                End If
                If kk!baa <> 0 Then
                    Print #1, Tab(T5); "BY BANK "; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                    netamount = netamount - Round(kk!baa, 2)
                End If
           End If
           Print #1, Tab(T6); repli("-", 22)
           Print #1, Tab(T5); Chr(71) + "BALANCE : "; Tab(T8 + 5); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12) + Chr(72)
        
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
            Print #1, Tab(LEFTM); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!CNAME)) * 0.75))); "FOR " + Trim(tempdata!CNAME)
            Print #1, Tab(LEFTM); ""

       'PRINT THE FOOTER IN INVOICE END
       
       
        
        
        
        
        
        Close #1
        PrintOption.Show
        'Me.Enabled = False
        'viewinvoice.Left = 0
        'viewinvoice.Top = 10
        'viewinvoice.Show



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
kkk.Open "select * from setup1", CON, adOpenDynamic, adLockReadOnly, adCmdText
If Not kkk.BOF Then
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, Chr(27) + Chr(71); Chr(27) + Chr(15) + Chr(14)
      Print #1, Tab(((paperWidth - (Len(Trim(kkk!CNAME)) * 2)) / 2)); Chr(27) + Chr(15) + Chr(14); Trim(kkk!CNAME)
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
rs1.Open "CASHA", CON, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To,"; Tab(T1 - 3); IIf(Optioncash.value = True, rs1!CASHPARTYNAME, rs1!SUBLEDGER); Tab(T5); "Cash Memo No. : "; Trim(rs1!INVOICENO); Tab(T8 + 5); "Dt. :"; Tab(T8 + 12); rs1!invoicedate 'Chr(27) + Chr(15);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); kkk!DESCFORINVOICE
        Print #1, Tab(3); kkk!ADDRESS1; Tab(T5); "Order by     : "; Trim(rs1!ORDERBY); Tab(T8 + 5); "Dt. :"; Tab(T8 + 12); rs1!ORDERDATE
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
    kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
                Print #1, rsets(Trim(str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(str(kk!quantity)), 5); Tab(T6 + 1); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!quantity
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
                tdata.Open "select sum(amount) from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                    Print #1, Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0), 2)), "0.00")), 12)
                    Print #1, Tab(T5); "Less Discount @ " + Trim(Format(str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
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
        Print #1, Tab(T5 - 4); rsets(Trim(str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12)
        
        Line = Line + 2
        
        
        If kk.State = 1 Then
             kk.close
        End If
        kk.Open "Select * from CASHC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
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
                        Print #1, Tab(T5); Trim(kk!Text) + "    " + Trim(Format(str(Round(kk!rate, 2)), "0.00")); Tab(T8 + 5); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5); Trim(kk!Text); Tab(T8 + 5); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T8); repli("-", 22)
            Print #1, Chr(27) + Chr(71); Tab(T5); "NET AMOUNT: "; Tab(T8 + 5); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
            VNetamt = netamount
            Line = Line + 2
        End If
        kk.close
        kk.Open "Select * from CASHA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(T5); kk!txt1; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + Round(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(T5); kk!txt2; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + Round(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(T5); "CASH RECD. "; Tab(T8 + 4); rsets(Trim(Format(str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - Round(kk!baa, 2)
             End If
        End If
        Print #1, Tab(T8); repli("-", 22)
        Print #1, Tab(T5); Chr(27) + Chr(71); "BALANCE : "; Tab(T8 + 5); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
        Line = Line + 2
        'PRINT THE FOOTER IN INVOICE START
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
        tempdata.Open "setup1", CON, adOpenDynamic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!CNAME)) * 0.65))); "FOR " + Trim(tempdata!CNAME)
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
    kkk.Open "select * from setup1", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
        'Print #1, Tab(1); "E.& O.E"
        'Print #1, Tab(1); kkk!COURT; Tab(LEFTM + (paperWidth - ((Len(kkk!COURT) + Len(kkk!Cname)) * 0.75))); "FOR " + Trim(kkk!Cname)
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); kkk!COURT; Tab(75); "FOR " + Trim(kkk!CNAME)
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
     Print #1, Tab(((paperWidth - (Len(Trim(kkk!CNAME)) * 2)) / 2)); Chr(27) + Chr(77) + Chr(14); Trim(kkk!CNAME)
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
rs1.Open "CASHA", CON, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
'Tab(20); Mid$(rs1!SUBLEDGER, 1, 5);
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); " To,"; Tab(7); IIf(Optioncash.value = True, "", Mid$(rs1!SUBLEDGER, 1, 5)); Tab(48); "Cash Memo No. : "; Trim(rs1!INVOICENO); Tab(82); "Dt. : "; rs1!invoicedate; Chr(27) + Chr(72);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & IIf(Optioncash.value = True, rs1!CASHPARTYNAME, kkk!DESCFORINVOICE)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS1), "", kkk!ADDRESS1); Tab(49); Chr(27) + Chr(71); "Order by    : "; Chr(27) + Chr(72); Trim(rs1!ORDERBY); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS2), "", kkk!ADDRESS2); Tab(49); Chr(27) + Chr(71); "Bilty No.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS3), "", kkk!ADDRESS3)
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
    kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,sno ", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
                'Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(Str(kk!quantity)), 5); Tab(T6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                Print #1, Tab(0); rsets(Trim(str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(str(kk!quantity)), 5); Tab(58); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!quantity
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CashB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(str(tdata(0)), "0.00")), 12)
                   'Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(Str(tdata(0) * cdiscount / 100), "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata(0) - (tdata(0) * cdiscount / 100)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(70); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.close
             Loop
         End If
    End If
    Print #1, repli("-", 96)
    Print #1, Tab(52); rsets(Trim(str(totalquantity)), 5); Tab(84); rsets(Trim(Format(str(netamount), "0.00")), 12)
    Line = Line + 2
    If kk.State = 1 Then kk.close
    kk.Open "Select * from CASHC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
       Do While Not kk.EOF
             If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(60); Trim(kk!Text) + "    " + Trim(Format(str(kk!rate), "0.00")); Tab(84); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(60); Trim(kk!Text); Tab(84); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
          
        End If
        Print #1, Tab(84); repli("-", 12)
        Print #1, Chr(27) + Chr(71); Tab(60); "NET AMOUNT: "; Tab(85); rsets(Trim(Format(str(netamount), "0.00")), 12); Chr(27) + Chr(72)
        VNetamt = netamount
        Line = Line + 2
        kk.close
        kk.Open "Select * from CASHA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(60); kk!txt1 & "    :"; Tab(84); rsets(Trim(Format(str(Abs(kk!txt1a)), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + kk!txt1a
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(60); kk!txt2 & " :"; Tab(84); rsets(Trim(Format(str(Abs(kk!txt2a)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + kk!txt2a
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(59); "CASH RECD. :"; Tab(84); rsets(Trim(Format(str(Abs(kk!baa)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - kk!baa
             End If
             If netamount <> 0 Then
                 Print #1, Tab(84); repli("-", 12)
                 Print #1, Tab(59); Chr(27) + Chr(71); "BALANCE   : "; Tab(85); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
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
        tempdata.Open "setup1", CON, adOpenDynamic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(65); "FOR " + Trim(tempdata!CNAME)
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
