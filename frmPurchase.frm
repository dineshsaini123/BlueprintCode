VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPurchase 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11715
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OPTFREIGHTPAID 
      BackColor       =   &H8000000B&
      Caption         =   "FREIGHT PAID"
      Height          =   315
      Left            =   7050
      TabIndex        =   70
      Top             =   5730
      Width           =   1455
   End
   Begin VB.ComboBox customercode 
      Height          =   960
      Left            =   7080
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Top             =   390
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.ComboBox Genledger 
      Height          =   315
      Left            =   1050
      Sorted          =   -1  'True
      TabIndex        =   39
      Top             =   6930
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Picture5 
      Height          =   510
      Left            =   1380
      ScaleHeight     =   450
      ScaleWidth      =   9600
      TabIndex        =   32
      Top             =   6750
      Width           =   9660
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
         Height          =   420
         Left            =   7530
         TabIndex        =   69
         Top             =   30
         Width           =   1020
      End
      Begin VB.CommandButton Commandother 
         Caption         =   "&Other"
         Height          =   450
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   930
      End
      Begin VB.CommandButton Commandhelp 
         Caption         =   "Help"
         Height          =   255
         Left            =   -300
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton Commandedit 
         Caption         =   "&Edit"
         Height          =   450
         Left            =   2100
         TabIndex        =   19
         Top             =   0
         Width           =   1140
      End
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Height          =   450
         Left            =   3240
         TabIndex        =   20
         Top             =   0
         Width           =   1140
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   450
         Left            =   4380
         TabIndex        =   21
         Top             =   0
         Width           =   990
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   450
         Left            =   5385
         TabIndex        =   22
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   450
         Left            =   6450
         TabIndex        =   23
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   6645
         TabIndex        =   34
         Top             =   600
         Width           =   915
      End
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   420
         Left            =   8580
         TabIndex        =   24
         Top             =   30
         Width           =   990
      End
      Begin VB.CommandButton Commandadd 
         Caption         =   "&Add"
         Height          =   450
         Left            =   945
         TabIndex        =   18
         Top             =   0
         Width           =   1140
      End
      Begin VB.CommandButton Commandprintnh 
         Caption         =   "N&HPrint"
         Height          =   375
         Left            =   5715
         TabIndex        =   33
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.ComboBox Bookcode 
      Height          =   2325
      ItemData        =   "frmPurchase.frx":0000
      Left            =   2880
      List            =   "frmPurchase.frx":0002
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   31
      Top             =   1935
      Width           =   2310
   End
   Begin VB.ComboBox Bookname 
      Height          =   2325
      Left            =   5220
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   30
      Top             =   1935
      Width           =   2295
   End
   Begin VB.CommandButton Commandall 
      Caption         =   "All Books"
      Height          =   375
      Left            =   -195
      TabIndex        =   29
      Top             =   5985
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.ComboBox cmbAgentName 
      Height          =   315
      Left            =   7110
      TabIndex        =   8
      Top             =   765
      Width           =   4500
   End
   Begin VB.TextBox txtadst 
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Top             =   6990
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   11145
      Top             =   7395
   End
   Begin VB.TextBox orno 
      Height          =   315
      Left            =   2220
      TabIndex        =   3
      Top             =   705
      Width           =   1200
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   300
      Top             =   5775
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox textbox 
      Height          =   315
      Left            =   7095
      TabIndex        =   6
      Top             =   405
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3795
      Left            =   165
      TabIndex        =   16
      Top             =   1755
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   6694
      _Version        =   393216
      FillStyle       =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin MSMask.MaskEdBox through 
      Height          =   285
      Left            =   -45
      TabIndex        =   26
      Top             =   1980
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox I_DTOB 
      Height          =   315
      Left            =   5055
      TabIndex        =   5
      Top             =   705
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bundles 
      Height          =   285
      Left            =   1950
      TabIndex        =   25
      Top             =   6030
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox I_OB 
      Height          =   315
      Left            =   3420
      TabIndex        =   4
      Top             =   705
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
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
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   705
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tempmeb 
      Height          =   285
      Left            =   225
      TabIndex        =   36
      Top             =   2460
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
      Left            =   1125
      TabIndex        =   37
      Top             =   4005
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
      Left            =   165
      TabIndex        =   38
      Top             =   4395
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
      Left            =   180
      TabIndex        =   0
      Top             =   705
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox through1 
      Height          =   285
      Left            =   7890
      TabIndex        =   15
      Top             =   1380
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox marka 
      Height          =   315
      Left            =   2220
      TabIndex        =   11
      Top             =   1365
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox weight 
      Height          =   315
      Left            =   3210
      TabIndex        =   12
      Top             =   1365
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox freight 
      Height          =   315
      Left            =   -45
      TabIndex        =   28
      Top             =   1980
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bdated 
      Height          =   315
      Left            =   6840
      TabIndex        =   14
      Top             =   1350
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox biltno 
      Height          =   315
      Left            =   5790
      TabIndex        =   13
      Top             =   1365
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox station 
      Height          =   315
      Left            =   -75
      TabIndex        =   27
      Top             =   1995
      Visible         =   0   'False
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtbilldt 
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
      Left            =   1200
      TabIndex        =   10
      Top             =   1365
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtbillno 
      Height          =   315
      Left            =   180
      TabIndex        =   9
      Top             =   1365
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   1200
      TabIndex        =   68
      Top             =   1050
      Width           =   1020
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bill No. : "
      Height          =   285
      Left            =   180
      TabIndex        =   67
      Top             =   1050
      Width           =   1005
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Discount : "
      Height          =   255
      Left            =   8760
      TabIndex        =   66
      Top             =   6060
      Width           =   1260
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Weight : "
      Height          =   300
      Left            =   3240
      TabIndex        =   65
      Top             =   1050
      Width           =   2535
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Freight : "
      Height          =   285
      Left            =   780
      TabIndex        =   64
      Top             =   5670
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bilty No. : "
      Height          =   285
      Left            =   5790
      TabIndex        =   63
      Top             =   1050
      Width           =   1005
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Railway/Station : "
      Height          =   285
      Left            =   2010
      TabIndex        =   62
      Top             =   5700
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bale(s) : "
      Height          =   285
      Left            =   1770
      TabIndex        =   61
      Top             =   6390
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   5055
      TabIndex        =   60
      Top             =   405
      Width           =   1020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Order By : "
      Height          =   285
      Left            =   3420
      TabIndex        =   59
      Top             =   405
      Width           =   1635
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Gross Amount : "
      Height          =   255
      Left            =   8775
      TabIndex        =   58
      Top             =   5745
      Width           =   1260
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Net Amount : "
      Height          =   255
      Left            =   8760
      TabIndex        =   57
      Top             =   6375
      Width           =   1200
   End
   Begin VB.Label label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cust Code : "
      Height          =   285
      Left            =   6120
      TabIndex        =   56
      Top             =   405
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Invoice No. : "
      Height          =   285
      Left            =   180
      TabIndex        =   55
      Top             =   405
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   1200
      TabIndex        =   54
      Top             =   405
      Width           =   1020
   End
   Begin VB.Label mga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   10080
      TabIndex        =   53
      Top             =   5745
      Width           =   1200
   End
   Begin VB.Label mna 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   10080
      TabIndex        =   52
      Top             =   6390
      Width           =   1200
   End
   Begin VB.Label mgd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   10080
      TabIndex        =   51
      Top             =   6045
      Width           =   1200
   End
   Begin VB.Label tqu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   5100
      TabIndex        =   50
      Top             =   5730
      Width           =   1155
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Quantity : "
      Height          =   255
      Left            =   3720
      TabIndex        =   49
      Top             =   5730
      Width           =   1350
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marka : "
      Height          =   285
      Left            =   2220
      TabIndex        =   48
      Top             =   1050
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   6840
      TabIndex        =   47
      Top             =   1050
      Width           =   1005
   End
   Begin VB.Label labelbybank 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2175
      TabIndex        =   46
      Top             =   6750
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label labelbybanklbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By Bank : "
      Height          =   255
      Left            =   2400
      TabIndex        =   45
      Top             =   6750
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label4 
      Caption         =   "Press F4 Key To Delete A Invoive Item"
      Height          =   405
      Left            =   5025
      TabIndex        =   44
      Top             =   7455
      Width           =   3705
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agent :"
      Height          =   315
      Left            =   6120
      TabIndex        =   43
      Top             =   705
      Width           =   960
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Through : "
      Height          =   285
      Left            =   7860
      TabIndex        =   42
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Purchase Entry :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   180
      TabIndex        =   41
      Top             =   -60
      Width           =   1995
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Order No : "
      Height          =   270
      Left            =   2220
      TabIndex        =   40
      Top             =   405
      Width           =   1200
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kk As ADODB.Recordset
'Dim CON As ADODB.Connection
Dim i As Integer
Dim lastrow, lastcol As Integer
Dim VALIDRATE As Boolean
Dim maxrow As Integer
Public totalamount, totaldiscount As Double
Public otheramount, otherdiscount As Double
Dim autoscroll As Boolean
Public edit As Boolean
Dim addmode As Boolean
Dim printheader As Boolean
Dim addoredit As Boolean
Dim inv As Integer
Dim inv1 As ADODB.Recordset

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
Dim rs As ADODB.Recordset
Dim Pno As Integer
Set rs = New ADODB.Recordset
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
Open "" + App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
FooterYes = False
header:
    If kkk.State = 1 Then
          kkk.Close
    End If
    CNSetup
    kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
        Print #1, Tab(1); kkk!COURT; Tab(75); "FOR " + Trim(kkk!cname)
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
If printheader = True Then
   If Not kkk.BOF Then
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, Chr(27) + Chr(71); Chr(27) + Chr(77) + Chr(14)
     Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15); Chr(27) + Chr(77) + Chr(14); dspace(Trim(kkk!cname))
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

Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("I N V O I C E")))) / 2 - 3); Chr(14); "I N V O I C E"; Chr(20); Tab(52); IIf(printheader = True, kkk!uptt, "")
Line = Line + 1
If printheader = True Then
   Print #1, Tab(63); kkk!cst
   Line = Line + 1
End If
If printheader = False Then
   Print #1, ""
   Line = Line + 1
End If
Print #1, repli("-", 96)
Line = Line + 1
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from Purchasea where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); "To,   S.L. Code : "; Tab(20); Mid$(rs1!subledger, 1, 5); Tab(50); "Invoice No. : "; Chr(27) + Chr(72); Trim(rs1!INVOICENO); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); rs1!InvoiceDate
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & kkk!DESCFORINVOICE
        Print #1, Tab(5); IIf(IsNull(kkk!address1), " ", kkk!address1); Tab(49); Chr(27) + Chr(71); "Order by    : "; Chr(27) + Chr(72); Trim(rs1!ORDERBY); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!OrderDate), "  /  /    ", rs1!OrderDate)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS2), " ", kkk!ADDRESS2)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS3), " ", kkk!ADDRESS3); Tab(49); Chr(27) + Chr(71); "Bilty No.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, ""
        kkk.Close
        Print #1, Chr(27) + Chr(71); "Through  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!through) + IIf(Trim(rs1!through1) = "", "", "," & rs1!through1)
        Print #1, Chr(27) + Chr(71); "Station  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!station); Tab(71); Chr(27) + Chr(71); "Pvt. Mark   : "; Chr(27) + Chr(72); Trim(rs1!marka)
        Print #1, Chr(27) + Chr(71); "Freight  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!freight); Tab(40); Chr(27) + Chr(71); "Weight    : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."; Tab(73); Chr(27) + Chr(71); "Bundle(s)   : "; Chr(27) + Chr(72); Trim(rs1!bundles)
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
    kk.Open "select * from Purchaseb where invoiceno=" + Trim(rs1!INVOICENO) + " and  " & stringyear & "   order by discount,sno ", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "' ", CON, adOpenKeyset, adLockReadOnly, adCmdText
                Print #1, Tab(0); rsets(Trim(str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(str(kk!quantity)), 5); Tab(58); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!quantity
                Line = Line + 1
                If Line > MaxLine - 6 Then
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from Purchaseb where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " and  " & stringyear & "   group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(str(tdata(0)), "0.00")), 12)
                   'Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(Str(tdata(0) * cdiscount / 100), "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata(0) - (tdata(0) * cdiscount / 100)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(70); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.Close
             Loop
           End If
       End If
       Print #1, repli("-", 96)
       Print #1, Tab(52); rsets(Trim(str(totalquantity)), 5); Tab(84); rsets(Trim(Format(str(netamount), "0.00")), 12)
       Line = Line + 2
       If kk.State = 1 Then
             kk.Close
       End If
       kk.Open "Select * from Purchasec where invoiceno=" + Trim(Me.I_NO.Text) & " and " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
       If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DEBITORCREDIT) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(60); Trim(kk!Text) + " :  @  " + Trim(Format(str(kk!rate), "0.00")) & " % "; Tab(84); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(60); Trim(kk!Text) & " :"; Tab(84); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
        End If
        Print #1, Tab(84); repli("-", 12)
        Print #1, Chr(27) + Chr(71); Tab(46); "NET AMOUNT  : "; Tab(85); rsets(Trim(Format(str(netamount), "0.00")), 12); Chr(27) + Chr(72)
        Print #1, Tab(84); repli("-", 12)
        VNetamt = netamount
        Line = Line + 3
        kk.Close
        Dim Va As Variant
        kk.Open "Select * from Purchasea where invoiceno=" + Trim(Me.I_NO.Text) & " and " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kk.BOF Then
             If kk!txt1a <> 0 Then
                Print #1, Tab(60); kk!txt1 & "  :"; Tab(84); rsets(Trim(Format(str(Abs(kk!txt1a)), "0.00")), 12)
                Line = Line + 1
                Va = netamount
                Va = Va + kk!txt1a
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(60); kk!txt2 & " :"; Tab(84); rsets(Trim(Format(str(Abs(kk!txt2a)), "0.00")), 12)
                 Line = Line + 1
                 Va = netamount
                 Va = Va - kk!txt2a
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(45); "BY BANK     :"; Tab(84); rsets(Trim(Format(str(Abs(kk!baa)), "0.00")), 12)
                 Line = Line + 1
                 Va = netamount
                 Va = netamount - kk!baa
             End If
        
            If kk!baa <> 0 Then
               Print #1, Tab(84); repli("-", 12)
               Print #1, Tab(45); "BALANCE     : "; Tab(84); rsets(Trim(Format(str(Va), "0.00")), 12);
               Print #1, Tab(84); repli("-", 12)
               Line = Line + 3
            End If
        End If
       'PRINT THE FOOTER IN INVOICE START
        Do While Line < 61
            Print #1, ""
            Line = Line + 1
        Loop
        Print #1, Tab(0); Chr(27) + Chr(71); toword(myround(VNetamt, 2)); Chr(27) + Chr(72)
        Print #1, Tab(0); repli("-", 96)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        CNSetup
        tempdata.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(75); "FOR " + Trim(tempdata!cname)
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
'frmPurchaseOther.calc
     mga.Caption = Format(myround(totalamount, 2), "0.00")
  '   mgd.Caption = Format(myround(totaldiscount, 2), "0.00")
     'mna.Caption = Format(myround((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
     mna.Caption = Format(myround((totalamount + otheramount - otherdiscount), 0), "0.00")
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
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If kk.State = 1 Then
   kk.Close
End If
If edit = False Then
  '  kk.Open "select max(invoiceno) from Purchasea", con, adOpenStatic, adLockReadOnly, adCmdText
   ' If kk(0) > 0 Then
   '     kk.MoveLast
   '     Me.I_NO.Text = Trim(Str(kk(0) + 1))
   ' Else
    '    Me.I_NO.Text = "1"
    'End If
    'kk.Close
    End If
        Dim ctl As Control
        For Each ctl In Me.Controls
            If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
                If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) And UCase(Trim(ctl.Name)) <> UCase(Trim("txtbilldt")) Then
                    ctl.Text = ""
                End If
                ctl.Enabled = False
            End If
        Next
        For i = 1 To maxrow
           Grid1.row = i
            For J = 1 To 10
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
        Unload frmPurchaseOther
End Sub
Public Function templost() As Boolean
    Dim check As Boolean
    Dim Check1 As Boolean
    Dim row, col As Integer
    Dim RRR, CCC As Integer
    Dim r, q, D As Double
    Dim mprevcol As Integer
    Dim mq As Currency, mr As Currency, mrot As Currency
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
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
                If rs.State = 1 Then
                    rs.Close
                End If
                rs.Open "select * from books where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
                row = Grid1.row
                col = Grid1.col
                If Trim(Grid1.Text) <> "" Then
                    If Not rs.BOF Then
                        rs.MoveFirst
                        rs.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        'rs.Find "bookcode='" + Trim(tempmeb.Text) + "'"
                        
                        If rs.EOF Then
                            tempmeb.Visible = True
                            tempmeb.SetFocus
                            rs.Close
                            templost = False
                            Exit Function
                        Else
                                    
                        'check1 = True
                            
                            Grid1.Text = rs(0)
                            Grid1.col = 2
                            Grid1.Text = rs(1) & ": " & rs!size1 & " " & rs!unit1 & " " & rs!size2 & " " & rs!unit2 & ": " & rs!quality
                            'Grid1.Text = rs(1)
                            
                         '   If Not edit Then
                                Grid1.col = 3
                                If Trim(Grid1.Text) = "" Then
                                    Grid1.Text = 0
                                End If
                                q = Val(Grid1.Text)
                                Grid1.col = 7
                                Grid1.Text = Grid1.row
  
                                Grid1.col = 8
                                If Trim(Grid1.Text) = "" Then
                                Grid1.Text = Format(rs(2), "0.00")            'rs(3)
                                r = rs(2)
                                
                                End If
                                                                
                                '/******************
                                
                                'Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
                                'Grid1.col = 6
                                'If Grid1.Text = "" Or addmode = True Then
                                'If Trim(kk(0)) <> "" Then
                                 '   tempstr = Trim(kk(0))
                                  '  kk.Close
                                   ' Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
                                   ' Grid1.col = 4
                                   ' If kk.BOF Then
                                    '    GoTo abc
                                    'End If
                                    'Grid1.Text = Format(kk(0), "0.00")
                                    'Grid1.col = 6
                                    'Grid1.Text = Format(kk(0), "0.00")
                                    'D = kk(0)
                                'Else
'abc:
                               '     Grid1.col = 4
                                '    Grid1.Text = Format(rs(4), "0.00")
                                 '   Grid1.col = 6
                                  '  Grid1.Text = Format(rs(4), "0.00")
                                  '  D = rs(4)
                                'End If
                            
                                                              
                                    Grid1.col = 9
                                    Grid1.Text = IIf(IsNull(rs!per), "", rs!per)
                                
                                Grid1.col = 10
                                Grid1.Text = Format(myround(q * r, 2), "0.00")
                           '     Grid1.col = 8
                           
                           ' Grid1.Text = Format(myround((q * r) * (D / 100), 2), "0.00")
                              'End If
                          '  End If
                            Grid1.col = col
                            rs.Close
                        End If
                    End If
                End If
            
            
            
            Case 3, 8
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
                Grid1.col = 8
                r = Val(Trim(Grid1.Text))
                Grid1.col = 10
                Grid1.Text = Format(myround(q * r, 2), "0.00")
                Grid1.col = col
''                Grid1.col = 5
''                r = Val(Trim(Grid1.Text))
''                Grid1.col = 6
''                D = Val(Trim(Grid1.Text))
''                Grid1.col = 7
''                Grid1.Text = Format(myround(q * r, 2), "0.00")
''                Grid1.col = 8
''                Grid1.Text = Format(myround((q * r) * (D / 100), 2), "0.00")
''                Grid1.col = col
             Case 4
                Grid1.Text = tempmeb.Text
                If Trim(Grid1.Text) = "" Then
                    Grid1.Text = 0
                End If
                
                Case 5, 6
                Grid1.Text = tempmeb.Text
                
                Case 7, 8, 9
                'tempmeb.Text = Grid1.Text
                Grid1.Text = tempmeb.Text
                                
        End Select
        row = Grid1.row
        col = Grid1.col
        totalamount = 0
        totaldiscount = 0
        For i = 1 To maxrow
            Grid1.row = i
            Grid1.col = 10
            totalamount = totalamount + myround(Val(Trim(Grid1.Text)), 2)
            '' Grid1.col = 7
''            totalamount = totalamount +myround(Val(Trim(Grid1.Text)), 2)
''            Grid1.col = 8
''            totaldiscount = totaldiscount +myround(Val(Trim(Grid1.Text)), 2)
        Next
        invoicecalc
        Me.tqu.Caption = ""
        For i = 1 To maxrow
            Grid1.col = 3
            Grid1.row = i
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
        Next
        
        Grid1.row = RRR
        'If check1 = True Then
        'Grid1.col = 3
        'Else
        Grid1.col = CCC
        'End If
        'check1 = False
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





Private Sub Bookname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        
        Dim mprevcol As Integer
        Dim mq As Currency, mr As Currency, mrot As Currency
        mprevcol = Grid1.col
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Select Case Grid1.col
            
            Case 2
                Dim row, col As Integer
                row = Grid1.row
                col = Grid1.col
                If Trim(Bookname.Text) = "" Then
                    Grid1.col = 1
                    If Trim(Grid1.Text) = "" Then
                        Grid1.Text = Bookname.Text
                           'Bookname.SetFocus
'********* vk
                          If Trim(Grid1.Text) = "" And row = 1 Then
                                 Grid1.col = 2
                                 Grid1.Text = ""
                                 If Trim(Grid1.Text) = "" Then
                                          Grid1.col = 1
                                          Bookname.SetFocus
                                          Grid1.SetFocus
                                          Grid1_Click
                                       Exit Sub
                                 End If
                            Else
                                 Commandother.SetFocus
                           End If
'********           'Commandother.SetFocus
                           Exit Sub
                    End If
                End If
                Grid1.row = row
                Grid1.col = col
                Grid1.Text = Bookname.Text
                '/*************************
                If rs.State = 1 Then
                    rs.Close
                End If
                rs.Open "select * from books where " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
                row = Grid1.row
                col = Grid1.col
                If Trim(Grid1.Text) <> "" Then
                    If Not rs.BOF Then
                        rs.MoveFirst
                        
                        
                        rs.Find "BOOKCODE='" + Trim(VBA.Left(Grid1.Text, InStr(Grid1.Text, "-") - 1)) + "'"
                        'rs.Find "bookname='" + Trim(VBA.Left(Grid1.Text, InStr(Grid1.Text, ":") - 1)) + "'"
                        If rs.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            rs.Close
                            Exit Sub
                        Else
                            Grid1.col = 1
                            Grid1.Text = rs(0)
                            Grid1.col = 2
                            
                            Grid1.Text = rs(1) & ": " & rs!size1 & " " & rs!unit1 & " " & rs!size2 & " " & rs!unit2 & ": " & rs!quality
                            Grid1.col = 3
                            If Trim(Grid1.Text) = "" Then
                                Grid1.Text = 0
                            End If
                            q = Val(Grid1.Text)
                            
                            Grid1.col = 7
                            Grid1.Text = Grid1.row
                                                        
                            Grid1.col = 8
                            Grid1.Text = Format(rs(2), "0.00")
                            r = rs(2)
                                
               Grid1.col = 9
               Grid1.Text = IIf(IsNull(rs!per), "", rs!per)
                                Grid1.col = 10
                                Grid1.Text = myround(q * r, 2)
                            Grid1.col = col
                            rs.Close
                        End If
                    End If
                End If
        End Select
        
        row = Grid1.row
        col = Grid1.col
        totalamount = 0
        totaldiscount = 0
        For i = 1 To maxrow
            Grid1.row = i
            Grid1.col = 10
            totalamount = totalamount + myround(Val(Trim(Grid1.Text)), 2)
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
            Case 3, 4, 5, 7, 8, 9    'R 7,8,9
                Grid1.col = Grid1.col + 1
                Grid1.SetFocus
                Grid1_Click
            Case 9 'R 6 replace with 9
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
'If cmbAgentName.Text = "" Then
'   'MsgBox "Enter a Agent Name.. "
'   'cmbAgentName.SetFocus
'   'Exit Sub
'Else
'  Dim rs1 As New ADODB.Recordset
'  rs1.Open "select * from agentmaster where " & stringyear & " and agentname='" & cmbAgentName.Text & "' order by agentname", CON, adOpenKeyset, adLockReadOnly, adCmdText
'  If rs1.RecordCount <= 0 Then
'     MsgBox "Enter valid Agent Name.. "
'     'cmbAgentName.SetFocus
'  End If
'End If

End Sub

Private Sub cmdprint_Click()
    MainMenu.cr1.SelectionFormula = ""
    MainMenu.cr1.Connect = constr
    MainMenu.cr1.SelectionFormula = "{invoicea.invoiceno} = " & I_NO.Text & " AND {invoicea.setupid} = " & main.setupid & " AND {invoicea.fyear} = '" & main.session & "'"
    MainMenu.cr1.ReportFileName = strrptpath & "\reports\purchase.rpt"
    'MainMenu.cr1.Formulas(0) = "agentprint=" & IIf(chkagentprint.Value = 1, "'True'", "'False'")
    'MainMenu.cr1.Formulas(1) = "Manufacturedby=" & IIf(cbomanufacturedby.Text <> "", "'Manufactured By : " & cbomanufacturedby.Text & "'", "''")
    MainMenu.cr1.Action = 1
End Sub

Private Sub Commandprintnh_Click()
printheader = False
'printinvoice
End Sub

Private Sub Commandabandon_Click()
invoiceabandon
Set inv1 = New ADODB.Recordset

If inv1.State = 1 Then inv1.Close
inv1.Open "SELECT MAX(INVOICENO) FROM Purchasea where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
    If inv1.RecordCount > 0 Then
    'inv = IIf(IsNull(Trim(Str(inv1(0)))), 1, inv1(0))
        inv = IIf(IsNull(inv1(0)), 1, inv1(0))
    End If
    inv1.Close
Me.I_NO.Text = inv
Me.I_NO_LostFocus

Dim ctl As Control
        For Each ctl In frmPurchase.Controls
            If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is CrystalReport Then
                ctl.Enabled = False
            End If
            Next
    Picture5.Enabled = True
Me.Commandall.Enabled = False
'Me.Commandall.Enab1ed = False
Me.Commandother.Enabled = False
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

through.Text = "DIRECT"

End Sub
Private Sub Commandadd_Click()
    
    invoiceabandon
    through.Text = "DIRECT"
    Dim rs As ADODB.Recordset
    'inv = Val(Me.I_NO.Text)
    
    addoredit = True
    addmode = True
    Set rs = New ADODB.Recordset
    Dim TEMPNUM As Integer
    If edit = False Then
       If CON.Execute("Select max(invoiceno) from Purchasea where " & stringyear)(0) >= Val(Trim(Me.I_NO.Text)) Then
          
          Me.I_NO.Text = CON.Execute("Select max(invoiceno) from Purchasea where " & stringyear)(0) + 1
           inv = CON.Execute("Select max(invoiceno) from Purchasea where " & stringyear)(0)
          
          'rs.Open "select * from tempinv where  " & stringyear & "  ", CON, adOpenKeyset, adLockOptimistic, adcmdtext
          'If rs.BOF Then
           '  rs.AddNew
          'End If
          'rs!In = Val(Me.I_NO.Text)
          'rs.Update
          'rs.Close
          
     End If
    
    End If
    Dim ctl As Control
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is CrystalReport Then
                ctl.Enabled = True
        End If
        'If UCase(Trim(ctl.Name)) = UCase(Trim(Me.I_NO.Name)) Then
        '   ctl.Enabled = False
        'End If
    Next
    Me.edit = False
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
    i_dt.SetFocus
    
End Sub
Private Sub Commandall_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
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
    If rs.State = 1 Then
        rs.Close
    End If
    rs.Open "select * from books where  " & stringyear & "    order by BOOKCODE", CON, adOpenKeyset, adLockReadOnly, adCmdText
    row = Grid1.row
    col = Grid1.col
    If Not rs.BOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            Grid1.col = 1
            Grid1.Text = rs(0)
            Grid1.col = 2
            Grid1.Text = rs(1)
            Grid1.col = 3
            If Trim(Grid1.Text) = "" Then
                Grid1.Text = Val(myvalue)
            End If
            q = Val(Grid1.Text)
            Grid1.col = 5
            Grid1.Text = Format(rs(3), "0.00")            'rs(3)
            r = rs(3)
            '/******************
            Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "' and " & stringyear)
            Grid1.col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.Close
                Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "' and " & stringyear)
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
                Grid1.Text = Format(rs(4), "0.00")
                Grid1.col = 6
                Grid1.Text = Format(rs(4), "0.00")
                D = rs(4)
            End If
            Grid1.col = 7
            Grid1.Text = Format(myround(q * r, 2), "0.00")
            Grid1.col = 8
            Grid1.Text = Format(myround((q * r) * (D / 100), 2), "0.00")
            If Not rs.EOF Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.row = Grid1.row + 1
                rs.MoveNext
            End If
        Loop
        '/**fghfghgh
        '    Grid1.col = col
    End If
    rs.Close
   ' row = Grid1.row
   ' col = Grid1.col
    totalamount = 0
    totaldiscount = 0
    Me.tqu.Caption = ""
    For i = 1 To Grid1.Rows - 1
            Grid1.row = i
            Grid1.col = 7
            totalamount = totalamount + myround(Val(Trim(Grid1.Text)), 2)
            Grid1.col = 8
            totaldiscount = totaldiscount + myround(Val(Trim(Grid1.Text)), 2)
            Grid1.col = 3
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
     Next
     maxrow = Grid1.Rows - 1
Else
'Grid1_Click
Exit Sub
End If

invoicecalc

End Sub

Private Sub Commanddelete_Click()
If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
                CON.Execute ("DELETE from Purchasea where INVOICENO = " + Trim(I_NO.Text) & " and " & stringyear)
                CON.Execute ("DELETE from Purchaseb where INVOICENO = " + Trim(I_NO.Text) & " and " & stringyear)
                CON.Execute ("DELETE from Purchasec where INVOICENO = " + Trim(I_NO.Text) & " and " & stringyear)
                invoiceabandon
End If
End Sub
Private Sub Commandedit_Click()
    On Error Resume Next
    Commandadd.Enabled = False
    Me.Commandedit.Enabled = False
    Picture5.Enabled = True
    Commandadd.Enabled = False
    Commandedit.Enabled = False
    Commandall.Enabled = True
    Commandsave.Enabled = True
    Commanddelete.Enabled = False
    Commandsearch.Enabled = False
    CommandPrint.Enabled = False
     Commandprintnh.Enabled = False
    Grid1.Enabled = True
    Commandall.Enabled = False
    Me.customercode.Enabled = True
    edit = True
    addoredit = False
    I_NO_LostFocus
    i_dt.Enabled = True
    i_dt.SetFocus
    ' PurchaseTmp creation start
    CON.Execute ("DELETE from PurchaseTmp where " & stringyear)
    DoEvents
    CON.Execute ("insert into PurchaseTmp  select * from Purchasec where INVOICENO = " + Trim(I_NO.Text) & " and " & stringyear)
    DoEvents
    ' invoicetmp creation end
    Dim kx As Integer
    kx = 0
    Do While kx < 11000
    kx = kx + 1
    Loop
    addoredit = False
    
    HIT
    VB.Screen.ActiveForm.ActiveControl.SelLength = 2
    
    frmPurchaseOther.TOP = 0
    frmPurchaseOther.Left = 0
    frmPurchaseOther.Visible = False
End Sub
Private Sub Commandother_Click()
    Me.Enabled = False
    frmPurchaseOther.Show 1
    frmPurchaseOther.TOP = 0
    frmPurchaseOther.Left = 0
    
    'Unload frmPurchaseOther
    'Load frmPurchaseOther
    'frmPurchaseOther.Show
    'frmPurchaseOther.Top = 0
    'frmPurchaseOther.Left = 0
    'frmPurchaseOther.Show
  
 
End Sub
Private Sub CommandPrint_Click()
        
'cr1.GroupSelectionFormula = ""
        cr1.SelectionFormula = ""
        cr1.SelectionFormula = "{Purchasea.invoiceno} = " & frmPurchaseI_NO.Text & " and {purchasea. " & stringyear & "  }"
        cr1.ReportFileName = strrptpath & "\reports\frmPurchaserpt"
        cr1.WindowShowPrintBtn = True
        cr1.WindowShowPrintSetupBtn = True
        cr1.WindowState = crptMaximized
        cr1.Action = 1
        
  
''''  printheader = True
''''  printinvoice
   
End Sub
Private Sub Commandreturn_Click()
    Unload Me
    addoredit = False
    MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()
    
    Dim SAVED As Boolean
    Dim LAMOUNT As Double
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If MsgBox("Do you want to save it now ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    Grid1.row = 1
    Grid1.col = 1
    If Trim(Grid1.Text) = "" Then
       MsgBox "Please Enter item.... "
       Exit Sub
    End If
    SAVED = False
    If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(customercode.Text) <> "" Then
            
            If edit Then
                CON.Execute ("DELETE from Purchasea where INVOICENO = " + Trim(I_NO.Text) & " and " & stringyear)
                CON.Execute ("DELETE from Purchaseb where INVOICENO = " + Trim(I_NO.Text) & " and " & stringyear)
                CON.Execute ("DELETE from Purchasec where INVOICENO = " + Trim(I_NO.Text) & " and " & stringyear)
                CON.Execute "select * from PurchaseTmp where " & stringyear
            End If
            
            
            If rs.State = 1 Then
                rs.Close
            End If
            LAMOUNT = 0
            rs.Open "select * from Purchasea where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
            If Not edit Then
again:
           If CON.Execute("Select max(invoiceno) from Purchasea where " & stringyear)(0) >= Val(Trim(Me.I_NO.Text)) Then
                   ' Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'GoTo again
                End If
            End If
            
            rs.addNew
            rs!INVOICENO = Val(Me.I_NO.Text)
            rs!InvoiceDate = Me.i_dt.Text
            rs!Genledger = Trim(Me.Genledger.Text)
            'rs!subledger = Trim(Me.customercode.Text)
            rs!subledger = Trim(Me.textbox.Text)
            rs!agentname = Trim(Me.cmbAgentName.Text)
            rs!ORDERNO = Trim(Me.orno.Text)
            rs!ORDERBY = Trim(Me.I_OB.Text)
            If Trim(Me.I_DTOB) <> Trim("__/__/____") Then
            '    rs!ORDERDATE = Date
            'Else
                rs!OrderDate = Trim(Me.I_DTOB.Text)
            End If
            rs!billno = Val(Trim(Me.txtbillno.Text))
            If Trim(Me.txtbilldt) <> Trim("__/__/____") Then
                rs!BILLDT = Trim(Me.txtbilldt.Text)
            End If
            rs!marka = Trim(Me.marka.Text)
            rs!bundles = Trim(Me.bundles)
            rs!through = Trim(Me.through.Text)
            rs!through1 = Trim(Me.through1.Text)
            If Trim(Me.through1.Text) = "" Then
                rs!through1 = " "
            End If
            rs!station = Trim(Me.station.Text)
            rs!biltyno = Trim(Me.biltno.Text)
            If Trim(Me.bdated) <> Trim("__/__/____") Then
                rs!BILTYDATE = Me.bdated & ""
           End If
            rs!freight = Trim(Me.freight)
            rs!weight = Trim(Me.weight)
            rs!netamount = myround(Val(Trim(Me.mna.Caption)), 0)
            rs!gamount = (Me.totalamount - Me.totaldiscount)
            rs!txt1 = Trim(frmPurchaseOther.T1TEXT.Text)
            rs!txt1a = Val(Trim(frmPurchaseOther.T1.Text))
            rs!txt2 = Trim(frmPurchaseOther.T2TEXT.Text)
            rs!txt2a = Val(Trim(frmPurchaseOther.T2.Text))
            rs!baa = Val(Trim(frmPurchaseOther.T3TEXT.Text))
            rs!baa = Val(Trim(labelbybank.Caption))
            rs!freightpaid = OPTFREIGHTPAID.Value
            If addmode = True Then
            
            If Val(Trim(frmPurchaseOther.T3TEXT.Text)) <> 0 Then
               rs!advicestatus = "Pending"
            End If
            
            Else
               rs!advicestatus = Me.txtadst.Text & ""
            End If
            Dim trs As New ADODB.Recordset
            trs.Open " SELECT DISTCODE FROM SLEDGER  WHERE   " & stringyear & "  and SUBLEDGER='" & Me.textbox.Text & "'", CON, adOpenStatic, adLockOptimistic, adCmdText
            If Not trs.BOF Then
                rs!District = Trim(trs!distcode)
            Else
                rs!District = ""
            End If

                For i = 1 To frmPurchaseOther.Grid1.Rows - 1
                        If i = 1 Then
                        rs!aexp1 = frmPurchaseOther.Grid1.TextMatrix(1, 0)
                        rs!aexp1rate = Val(frmPurchaseOther.Grid1.TextMatrix(1, 1))
                        rs!aexp1am = Val(frmPurchaseOther.Grid1.TextMatrix(1, 2))
                        End If
                        
                        If i = 2 Then
                        rs!aexp2 = frmPurchaseOther.Grid1.TextMatrix(2, 0)
                        rs!aexp2rate = Val(frmPurchaseOther.Grid1.TextMatrix(2, 1))
                        rs!aexp2am = Val(frmPurchaseOther.Grid1.TextMatrix(2, 2))
                        End If
                        
                        If i = 3 Then
                        rs!aexp3 = frmPurchaseOther.Grid1.TextMatrix(3, 0)
                        rs!aexp3rate = Val(frmPurchaseOther.Grid1.TextMatrix(3, 1))
                        rs!aexp3am = Val(frmPurchaseOther.Grid1.TextMatrix(3, 2))
                        End If
                        
                        If i = 4 Then
                        rs!lexp1 = frmPurchaseOther.Grid1.TextMatrix(4, 0)
                        rs!lexp1rate = Val(Val(frmPurchaseOther.Grid1.TextMatrix(4, 1)))
                        rs!lexp1am = Val(frmPurchaseOther.Grid1.TextMatrix(4, 2))
                        End If
                        
                        If i = 5 Then
                        rs!lexp2 = frmPurchaseOther.Grid1.TextMatrix(5, 0)
                        rs!lexp2rate = Val(frmPurchaseOther.Grid1.TextMatrix(5, 1))
                        rs!lexp2am = Val(frmPurchaseOther.Grid1.TextMatrix(5, 2))
                        End If
                        
                        If i = 6 Then
                        rs!lexp3 = frmPurchaseOther.Grid1.TextMatrix(6, 0)
                        rs!lexp3rate = Val(frmPurchaseOther.Grid1.TextMatrix(6, 1))
                        rs!lexp3am = Val(frmPurchaseOther.Grid1.TextMatrix(6, 2))
                        End If
                        
                        'frmPurchaseOther.Grid1.row = frmPurchaseOther.Grid1.row + 1
                Next
            

err1:
           If Not edit Then
                If CON.Execute("Select max(invoiceno) from Purchasea where " & stringyear)(0) >= Val(Trim(Me.I_NO.Text)) Then
                    'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'rs!INVOICENO = Val(Me.I_NO.Text)
                    On Error GoTo err1
                End If
            End If
            rs!FYear = main.session: rs!setupid = main.setupid
            rs!createdby = main.username
            rs!createdon = Now
            ' " & stringyear & "
            rs.Update
            On Error GoTo 0
            rs.Close
            'If rs.State = 1 Then rs.Close
            rs.Open "select * from Purchaseb where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
            'rs!invoicedate = Me.i_dt.Text
            
            RRRR = Grid1.row
            CCCC = Grid1.col
            
            
            
            For i = 1 To maxrow
                Grid1.row = i
                Grid1.col = 1
                If Trim(Grid1.Text) <> "" Then
                    Grid1.col = 3
                    If Val(Trim(Grid1.Text)) > 0 Then
                       'Grid1.col = 5
                    '   If Val(Trim(Grid1.Text)) > 0 Then
                         rs.addNew
                         Grid1.col = 1
                         rs!INVOICENO = Val(Me.I_NO.Text)
                         rs!InvoiceDate = Me.i_dt.Text
                         rs!Genledger = Trim(Me.Genledger.Text)
                         rs!subledger = Trim(Me.textbox.Text)
                         rs!Bookcode = Trim(Grid1.Text)
                         Grid1.col = 3
                         rs!quantity = Trim(Grid1.Text)
                         Grid1.col = 4
                         rs!btno = Grid1.Text
                         Grid1.col = 5
                         rs!mgdate = Trim(Grid1.Text)
                         Grid1.col = 6
                         rs!expdate = Trim(Grid1.Text)
                         Grid1.col = 7
                         rs!PrintOrder = Trim(Grid1.Text)
                         Grid1.col = 8
                         rs!rate = Trim(Grid1.Text)
                         Grid1.col = 9
                         rs!per = Trim(Grid1.Text)
                         Grid1.col = 10
                         rs!amount = Trim(Grid1.Text)
                         'LAMOUNT = Val(Trim(Grid1.Text))
                         'Grid1.col = 4
                         'rs!printorder = Trim(Grid1.Text)
                         'Grid1.col = 6
                         'rs!discount = Trim(Grid1.Text)
                         'Grid1.col = 10
                         'rs!netamount = LAMOUNT - Trim(Grid1.Text)
                         rs!netamount = mna.Caption
                         'LAMOUNT = 0
                         rs!FYear = main.session: rs!setupid = main.setupid
                        rs!createdby = main.username
                        rs!createdon = Now
                         rs.Update
                     '  End If
                    End If
                End If
            Next
            rs.Close
            Grid1.TopRow = 1
            Grid1.row = 1
            Grid1.col = 1
            
            
            
            
            
            
            
            
            
            rs.Open "select * from Purchasec where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
            '/******
                'Dim I, x As Integer
                Dim temprs As ADODB.Recordset
                Set temprs = New ADODB.Recordset
                For i = 1 To frmPurchaseOther.mrow
                    frmPurchaseOther.Grid1.row = i
                    frmPurchaseOther.Grid1.col = 0
                    If Trim(frmPurchaseOther.Grid1.Text) <> "" Then
                        rs.addNew
                        rs!INVOICENO = Val(Me.I_NO.Text)
                        rs!InvoiceDate = Me.i_dt.Text
                        rs!gamount = (Me.totalamount - Me.totaldiscount)
                        rs!Text = Trim(frmPurchaseOther.Grid1.Text)
                        If temprs.State = 1 Then
                            temprs.Close
                        End If
                        If edit Then
                        temprs.Open "select * from PurchaseTmp where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
                        If frmPurchaseOther.Grid1.Text <> "" Then
                                temprs.Find "TEXT='" + Trim(frmPurchaseOther.Grid1.Text) + "'"
                                rs!Genledger = Trim(temprs!Genledger)
                                rs!subledger = Trim(temprs!subledger)
                                rs!DEBITORCREDIT = Trim(temprs!DEBITORCREDIT)
                                rs!RYN = temprs!RYN & ""
                        End If
                        temprs.Close
                        Else
                        
                        temprs.Open "select * from PurchaseEnd where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
                        If frmPurchaseOther.Grid1.Text <> "" Then
                                temprs.Find "TEXT='" + Trim(frmPurchaseOther.Grid1.Text) + "'"
                                rs!Genledger = Trim(temprs!Genledger)
                                rs!subledger = Trim(temprs!subledger)
                                rs!DEBITORCREDIT = Trim(temprs!DEBITORCREDIT)
                                rs!RYN = temprs!RYN & ""
                        End If
                        temprs.Close
                        End If
                        frmPurchaseOther.Grid1.col = 1
                        rs!rate = Val(Trim(frmPurchaseOther.Grid1.Text))
                        If Val(Trim(frmPurchaseOther.Grid1.Text)) > 0 Then
                            rs!amount = (Me.totalamount - Me.totaldiscount) * (Val(Trim(frmPurchaseOther.Grid1.Text)) / 100)
                        Else
                        frmPurchaseOther.Grid1.col = 2
                            rs!amount = Val(Trim(frmPurchaseOther.Grid1.Text))
                        End If
                        rs!FYear = main.session: rs!setupid = main.setupid
                        rs!createdby = main.username
                        rs!createdon = Now
                    rs.Update
                    End If
                Next
                rs.Close
                
                
                
''                rs.Open "select * from tempINV where " & stringyear, CON, adOpenDynamic, adLockOptimistic, adCmdText
''                If rs.BOF Then
''                    rs.AddNew
''                rs!createdby = main.username
''                rs!createdon = Now
''                Else
''                rs!updatedby = main.username
''                rs!updatedon = Now
''                End If
''                rs!In = CON.Execute("Select max(invoiceno) from Purchasea")(0)
''                rs!fyear = main.session: rs!setupid = main.setupid
''
''                rs.Update
''                rs.Close
                
                'If addmode = True Then
                 '   rs.Open "select * from tempinv where  " & stringyear & "  ", CON, adOpenKeyset, adLockOptimistic, adcmdtext
                  '  If rs.BOF Then
                   '     rs.AddNew
'                    End If
 '                   rs!In = Val(Me.I_NO.Text)
  '                  rs.Update
   '                 rs.Close
    '            End If
            SAVED = True
        End If
        If SAVED Then
            MsgBox "Record Saved"
            Unload frmPurchaseOther
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
   End Sub

Private Sub Commandsearch_Click()
    Unload frmPurchaseOther
    CON.Execute "select * from PurchaseTmp where " & stringyear
    Me.Enabled = False
    'searchscreen.Grid1.row = 0
    'searchscreen.Grid1.col = 0
    Call searchscreen.tempr(13, "Purchase")
End Sub

Private Sub customercode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Me.textbox.Text = Me.customercode.Text
   End If
End Sub

Private Sub customercode_LostFocus()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    
    rs.Open "select * from sledger where  " & stringyear & "   and gledger='SUNDRY CREDITORS' and subledger='" + Trim(customercode.Text) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If rs.RecordCount <= 0 Then
        customercode.SetFocus
        HIT
        rs.Close
        Exit Sub
    End If
    
    
    'Dim rs1 As ADODB.Recordset
    'Set rs1 = New ADODB.Recordset
    'If rs!distcode <> "" And addmode = True Then
     '  rs1.Open "Select * from Districts where  " & stringyear & " and Districtname = '" & rs!distcode & "'", CON, adOpenStatic, adLockReadOnly
      ' If rs1.RecordCount > 0 Then
       '   Me.cmbAgentName = IIf(IsNull(rs1!Agentname), "", rs1!Agentname)
       'End If
    'End If
    
    
    
    'rs.Close
    
'    Me.textbox.Text = Me.customercode.Text
'
    Me.customercode.Visible = False
    Me.customercode.Enabled = False
       
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      Dim vn As String
        'If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
                If Grid1.row > 1 Then
                    'vn = Grid1.TextMatrix(Grid1.RowSel, 1)
                    Grid1.RemoveItem Grid1.row

      'If Grid1.row >= 1 Then
      '     Grid1.RemoveItem
          ' Grid1.RemoveItem Grid1.row - 1
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
             If addmode = True Then
                SendKeys "{DOWN}"
             End If
             SendKeys "{TAB}"
        Else
            If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("grid1")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("tempmeb")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bookname")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("through1")) Then
                SendKeys ("{TAB}")
            End If
        End If
    End If
    
    
End Sub
Private Sub Form_Load()
  
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
    Grid1.Left = 90
    Grid1.TOP = 1900
   'Grid1.Top = 2500
    'Set CON = New ADODB.Connection
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Me.TOP = 0
    Me.Left = 0
    Grid1.Rows = 2
    Grid1.Cols = 1
    Grid1.Rows = 2
    Grid1.Cols = 11
    Grid1.row = 0
    Grid1.col = 1
    
    Grid1.Text = "Item Code "
    Grid1.col = Grid1.col + 1
    
    Grid1.Text = "Item Name"
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Quantity"
    
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Batch No"
    
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Mfd. Date"
    
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Expiry Date"
    
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Print. Ord."
    
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Rate"
    
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Per"
   
    
    Grid1.col = Grid1.col + 1
    Grid1.Text = "Amount"
    
    'Grid1.Text = "Disc %"
    'Grid1.col = Grid1.col + 1
    

    'Grid1.col = Grid1.col + 1
    'Grid1.Text = "Disc. Amount"
    
    Grid1.RowHeight(0) = Grid1.CellHeight + 50
    
    
    Grid1.ColWidth(0) = 100
    Grid1.ColWidth(1) = 750
    Grid1.ColWidth(2) = 4550
    Grid1.ColWidth(3) = 750
    Grid1.ColWidth(4) = 550
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 900
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 600
    Grid1.ColWidth(9) = 450
    Grid1.ColWidth(10) = 980
    
    Bookname.Height = 2325
    
    Me.CommandPrint.Enabled = True
    Me.Commandprintnh.Enabled = True
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from books where " & stringyear & " and GROUPCODE='No'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.Bookcode.AddItem rs(0)
            Me.Bookname.AddItem rs(0) & " - " & rs(1) & ": " & rs!size1 & " " & rs!unit1 & " " & rs!size2 & " " & rs!unit2 & ": " & rs!quality
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    
    
    Genledger.Text = "SUNDRY CREDITORS"
    rs.Open "select * from sledger where gledger='" + Trim(Genledger.Text) + "' and " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.customercode.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
     
     
     
     '*******Agent  combo fill
    
    rs.Open "select  Agentname from AgentMaster where  " & stringyear & " order by agentname", CON, adOpenKeyset, adLockReadOnly, adCmdText
    cmbAgentName.Clear
    If Not rs.EOF Then
       Do While Not rs.EOF
          If IsNull(rs(0)) = False Then
            Me.cmbAgentName.AddItem rs(0)
          End If
          If Not rs.EOF Then rs.MoveNext
        Loop
    End If
    rs.Close

    
    Bookcode.Left = Grid1.Left
    Bookcode.Visible = False
    Bookname.Visible = False
    Grid1.Rows = 100
    For i = 1 To 99
        Grid1.RowHeight(i) = 300
    Next
    Bookcode.Width = 1230
    Bookname.Width = 2830
    amount.Width = rate.Width
  
    
'    rs.Open "select * from tempinv where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
'    If Not rs.BOF Then
'        Me.I_NO.Text = rs!In
'        frmPurchase.Enabled = True
'        frmPurchase.edit = False
'        frmPurchase.I_NO_LostFocus
'        frmPurchase.I_NO.Enabled = False
'        lastrow = 0
'        lastcol = 1
'
'        Dim ctl As Control
'        For Each ctl In frmPurchase.Controls
'            If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is CrystalReport Then
'            ctl.Enabled = False
'            End If
'            If UCase(Trim(ctl.Name)) = UCase(Trim(frmPurchase.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(frmPurchase.Commandall.Name)) Then
'                ctl.Enabled = False
'            End If
'        Next
'
'
'        frmPurchase.Picture5.Enabled = True
'        addoredit = False
'        SendKeys "{TAB}"
'    Else
       
       kk.Open "SELECT MAX(INVOICENO) FROM Purchasea where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
       If kk(0) <> "" Then
           Me.I_NO.Text = Trim(str(kk(0) + 1))
       Else
          Me.I_NO.Text = "1"
       End If
       kk.Close
'    End If
'    rs.Close
    
    
       
    Commanddelete.Enabled = True
    Commandedit.Enabled = True
    Commandsave.Enabled = False
    lastrow = 0
    lastcol = 1
    'Dim ctl As Control
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is CrystalReport Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(Me.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    Picture5.Enabled = True
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete

Me.Label20.Enabled = True
Me.Timer1.Enabled = True
through.Text = "DIRECT"
Me.i_dt.Text = Date
End Sub

Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub

Private Sub Grid1_Click()

If Trim(Me.customercode.Text) <> "" Then

Dim PREVROW As Integer
Dim prevcol As Integer
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
prevcol = Grid1.col
PREVROW = Grid1.row

If Grid1.row > 1 Then
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
        Grid1.col = 1
        If prevcol > 1 And Trim(Grid1.Text) = "" Then
            Grid1.col = 2
            Grid1.SetFocus
            SendKeys Chr(13)
        Else
        'IF GRID1.COL
            Grid1.col = prevcol
            Grid1.SetFocus
            SendKeys Chr(13)
        End If
        'SendKeys Chr(13)
       
    End If
End If
End If

End Sub

Private Sub Grid1_GotFocus()

'If Grid1.col = 6 And Grid1.col = 5 Then
'tempmeb.Mask = "__/__/____"
'Else
'tempmeb.Mask = ""
'End If

End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)

If Trim(Me.customercode.Text) <> "" Then
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    If (KeyAscii = 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
        If mwritemode = addmode Or mwritemode = EditMode Then
            Dim mprevcol As Integer
            
            
            Select Case Grid1.col
            'Case 1, 3, 4, 7, 8, 9
            Case 1, 3, 4, 7, 8, 9, 5, 6
                    tempmeb.Mask = ""
                    tempmeb.MaxLength = 20
                
            'Case 5, 6
                    'tempmeb.Mask = "##/##/####"
                    
             '       tempmeb.MaxLength = 10
            Case 2
                    With Bookname
                        .Visible = True
                        .ZOrder
                    End With
                End Select

            
            Select Case Grid1.col
            
            Case 1, 3, 4, 5, 6, 7, 8, 9
                Bookname.Visible = False
                tempmeb.Visible = True: tempmeb.Enabled = True
                tempmeb.ZOrder
                
                If Grid1.col <> 1 Then
'                    If Grid1.col <> 3 Then
                    'If Grid1.col = 7 Then  'R
                    If Grid1.col = 8 Then  'R
                        tempmeb.Text = Format(Grid1.Text, "0.00")
                        'tempmeb.Text = Format(Grid1.Text, "0.00")
                      Else
                        'If Grid1.col = 5 Or Grid1.col = 6 Or Grid1.col = 9 Then 'R
                        'Else    'R
                        'tempmeb.Text = Format(Grid1.Text, "0")
                        'End If  'R
                        tempmeb.Text = Grid1.Text
                    End If
                Else
                    tempmeb.Text = Grid1.Text
                End If
                
                
                tempmeb.Width = Grid1.ColWidth(Grid1.col)
                tempmeb.Left = Grid1.CellLeft + 80
                tempmeb.TOP = Grid1.TOP + Grid1.CellTop '- 50
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.Text = Grid1.Text
                Bookname.TOP = Grid1.TOP + Grid1.CellTop
                Bookname.Left = Grid1.CellLeft + 80
                Bookname.Width = Grid1.ColWidth(Grid1.col)
            'Case 6
                
            Case Else
                Bookname.Visible = False
                tempmeb.Visible = False
            End Select
            
            
            
'''''            Select Case Grid1.col
'''''                Case 1, 3, 4, 7, 8, 9
'''''                    tempmeb.Mask = ""
'''''                    tempmeb.MaxLength = 20
'''''
'''''                Case 5, 6
'''''                    tempmeb.Mask = "##/##/####"
'''''                    tempmeb.MaxLength = 10
'''''
'''''                Case 2
'''''                    With Bookname
'''''                        .Visible = True
'''''                        .ZOrder
'''''                    End With
'''''                End Select
            
            
            Select Case Grid1.col
            Case 2
                Bookname.SetFocus
                If KeyAscii <> 13 Then
                    SendKeys Chr(KeyAscii)
                End If
            Case 1, 3, 4, 5, 6, 7, 8, 9
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
   PopupMenu dd, , Grid1.Left + X, Grid1.TOP + Y
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
If edit = True Then
    Commandother.Enabled = True
End If
End Sub
Private Sub i_dt_LostFocus()
If Trim(i_dt.Text) <> Trim("__/__/____") Then
    If Not checkdate(Trim(i_dt.Text), i_dt) Then
        i_dt.SetFocus
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
Dim rs As ADODB.Recordset

SetButton Commandadd, Commandedit, Commandsave, Commanddelete
Set rs = New ADODB.Recordset
    If Trim(I_NO.Text) = "" Then
        MsgBox "Invoice cannot be null"
        I_NO.SetFocus
    Else
        If rs.State = 1 Then
           rs.Close
        End If
        ''rs.Open "Purchasea", con, adOpenKeyset, adLockReadOnly, adcmdtext
        rs.Open "Select * from  Purchasea where INVOICENO = " + Trim(I_NO.Text) + " and " & stringyear, CON, adOpenStatic, adLockReadOnly
        ''rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
        
        If rs.EOF Then
            If addoredit = False Then
                MsgBox "Purchase not found"
                Exit Sub
            End If
            Exit Sub
        End If
        If addoredit Then
            MsgBox "Purchase already exist..."
            I_NO.SetFocus
            HIT
            Exit Sub
        End If
        
        Dim ctl As Control
        For Each ctl In Me.Controls
            If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is CrystalReport Then
                ctl.Enabled = True
            End If
        Next
        
        I_NO.Text = rs!INVOICENO
        Me.i_dt.Text = rs!InvoiceDate
        Me.Genledger.Text = Trim(rs!Genledger)
        Me.customercode.Text = Trim(rs!subledger)
        Me.cmbAgentName.Text = IIf(IsNull(rs!agentname), "", rs!agentname)
        
        Me.textbox.Text = Trim(rs!subledger)
        'Me.orno.Text = rs!ORDERNO
        Me.orno.Text = IIf(IsNull(rs!ORDERNO), "", Trim(rs!ORDERNO))
        Me.I_OB.Text = IIf(IsNull(rs!ORDERBY), "", Trim(rs!ORDERBY))
        If rs!OrderDate <> "" Then
        Me.I_DTOB.Text = rs!OrderDate
        Else
        Me.I_DTOB.Text = "__/__/____"
        End If
        Me.marka.Text = IIf(IsNull(rs!marka), "", Trim(rs!marka))
        Me.bundles = IIf(IsNull(rs!bundles), "", rs!bundles)
        Me.through.Text = IIf(IsNull(rs!through), "", rs!through)
        Me.through1.Text = IIf(IsNull(rs!through1), "", rs!through1)
        Me.station.Text = IIf(IsNull(rs!station), "", rs!station)
        Me.biltno.Text = IIf(IsNull(rs!biltyno), "", rs!biltyno)
        OPTFREIGHTPAID.Value = rs!freightpaid
       If rs!BILTYDATE <> "" Then
        Me.bdated = rs!BILTYDATE
        Else
        Me.bdated.Text = "__/__/____"
        End If
        Me.txtbillno.Text = IIf(IsNull(rs!billno), "", rs!billno)
        If rs!BILLDT <> "" Then
        Me.txtbilldt = rs!BILLDT
        Else
        Me.txtbilldt.Text = "__/__/____"
        End If
        
        Me.freight = IIf(IsNull(rs!freight), "", rs!freight)
        Me.weight = IIf(IsNull(rs!weight), "", rs!weight)
        'Me.labelbybank =myround(val(Trim(rs!baa)
        Me.labelbybank = Format(myround(rs!baa, 2), "0.00")
       ' mna.Caption = rs!netamount
        mna.Caption = Format(myround(rs!netamount, 0), "0.00")
        'Me.Combosldistrictcode.Text = IIf(IsNull(rs!district), "", rs!district)
        Me.txtadst = IIf(IsNull(rs!advicestatus), "", rs!advicestatus)
        rs.Close
       
       ' frmPurchaseOther.Form_Load
'*/**/*/*/*/*//*/*
        If rs.State = 1 Then
                rs.Close
        End If
'Commandedit.Enabled = True

       ' Unload frmPurchaseOther
        CON.Execute "select * from PurchaseTmp where " & stringyear
       ' frmPurchaseOther.Form_Load
       ' 'rs.Open "Purchaseb", con, adOpenKeyset, adLockReadOnly, adcmdtext
       rs.Open "Select * from Purchaseb where INVOICENO =" + Trim(I_NO.Text) + " and " & stringyear, CON, adOpenStatic, adLockReadOnly
       '' rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
       Grid1.TopRow = 2
       If Not rs.EOF Then
            Grid1.row = 1
            Grid1.col = 1
            Do While Not rs.EOF
               If Trim(rs!INVOICENO) = Trim(I_NO.Text) Then
                Grid1.col = 1
                Grid1.Text = Trim(rs!Bookcode)
                If kk.State = 1 Then
                    kk.Close
                End If
                kk.Open "select * from books where " & stringyear & " and bookcode='" + Trim(rs!Bookcode) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
                If kk.RecordCount > 0 Then
                Grid1.col = 2
                Grid1.Text = Trim(kk!Bookname & ": " & kk!size1 & " " & kk!unit1 & " " & kk!size2 & " " & kk!unit2 & ": " & kk!quality)
                Grid1.col = 3
                Grid1.Text = Trim(rs!quantity)
                Grid1.col = 4
                Grid1.Text = Trim(rs!btno)
                Grid1.col = 5
                Grid1.Text = Trim(rs!mgdate)
                Grid1.col = 6
                Grid1.Text = Trim(rs!expdate)
                Grid1.col = 7
                Grid1.Text = Trim(rs!PrintOrder)
                Grid1.col = 8
                Grid1.Text = Format(myround(rs!rate, 2), "0.00")
                Grid1.col = 9
                Grid1.Text = IIf(IsNull(Trim(rs!per)), "", (Trim(rs!per)))
                'Grid1.col = 4
                'Grid1.Text = Format(myround(rs!printorder, 2), "0.00")
                'Grid1.col = 6
                'Grid1.Text = Format(myround(rs!discount, 2), "0.00")
                'Grid1.col = 8
                'Grid1.Text = Format(myround(rs!amount * (rs!discount / 100), 2), "0.00")
                Grid1.col = 10
                Grid1.Text = Format(myround(rs!amount, 2), "0.00")
                End If
                End If
                If Not rs.EOF Then
                    rs.MoveNext
                End If
                Grid1.row = Grid1.row + 1
                Grid1.Rows = Grid1.Rows + 1
            Loop
            maxrow = Grid1.row
        '    Me.i_dt.SetFocus
        End If
        row = Grid1.row
        col = Grid1.col
        Grid1.TopRow = 1
        totalamount = 0
        totaldiscount = 0
        For i = 1 To maxrow
            Grid1.row = i
            Grid1.col = 10
            totalamount = totalamount + myround(Val(Trim(Grid1.Text)), 2)
            'Grid1.col = 10
            'totaldiscount = totaldiscount +myround(Val(Trim(Grid1.Text)), 2)
        Next
     mga.Caption = Format(myround(totalamount, 2), "0.00")
     'mgd.Caption = Format(myround(totaldiscount, 2), "0.00")
     Me.tqu.Caption = ""
        For i = 1 To maxrow
            Grid1.col = 3
            Grid1.row = i
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
        Next
        Grid1.row = RRR
        Grid1.col = CCC
       ' templost = True
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
'''''''If Grid1.col = 1 Or Grid1.col = 2 Then
'''''''        Grid1.Text = tempmeb.Text
'''''''Else
'''''''
''''''''    If Grid1.col = 3 Then
''''''''        Grid1.Text = Format(tempmeb.Text, "0")
''''''''    Else
''''''''        Grid1.Text = Format(tempmeb.Text, "0.00")
''''''''    End If
'''''''    If Grid1.col = 3 Or Grid1.col = 4 Then
'''''''                    Grid1.Text = Format(tempmeb.Text, "0")
'''''''    Else
'''''''
'''''''    'If Grid1.col = 5 And Grid1.Text = "" Then
'''''''    'tempmeb.Mask = "__/__/____"
'''''''    'Grid1.Text = tempmeb.Text
'''''''     '  End If
'''''''
'''''''    If Grid1.col = 5 Or Grid1.col = 6 Then
'''''''        Else
'''''''        Grid1.Text = Format(tempmeb.Text, "0.00")
'''''''
'''''''    End If
'''''''    End If
'''''''End If
End Sub
Private Sub tempmeb_GotFocus()
    HIT
    
'''''''    If Grid1.col = 5 And Grid1.Text = "" Then
'''''''    tempmeb.Text = Format(Now, "dd/mm/yyyy")
'''''''    End If
'''''''    If Grid1.col = 6 And Grid1.Text = "" Then
'''''''    tempmeb.Text = Format(Now + (365 * 2), "dd/mm/yyyy")
'''''''    End If



    'If Grid1.col = 6 And Grid1.Text = "" Then
    'Grid1.col = 5
    'If Grid1.Text <> "" Then
    'Grid1.col = 6
    'tempmeb.Text = Format(CDate(Grid1.TextMatrix(Grid1.row, 5)) + 2, "dd/mm/yyyy")
    'End If
    'End If
    
    
End Sub

Private Sub tempmeb_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        Dim rs As ADODB.Recordset
           Set rs = New ADODB.Recordset
            
            Select Case Grid1.col
                
                Case 1
                
                Grid1.Text = tempmeb.Text
                    If rs.State = 1 Then
                        rs.Close
                    End If
                    
                    rs.Open "select * from books where " & stringyear & " ", CON, adOpenStatic, adLockReadOnly, adCmdText
                    If Not rs.BOF Then
                        
                        rs.MoveFirst
                        rs.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If rs.EOF And Trim(Grid1.Text) <> "" Then
                            rs.Close
                            Exit Sub
                        Else
                          rs.Close
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
                    'Grid1.col = 3
                    Grid1_Click
                
                Case 3
                    If Val(tempmeb.Text) > 0 Then
                        'Grid1.col = Grid1.col + 2
                        Grid1.col = Grid1.col + 1
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                
                Case 4
'                    Grid1.col = Grid1.col + 2
                     Grid1.col = Grid1.col + 1
                    Grid1.SetFocus
                    Grid1_Click
                Case 5
                    '''''If tempmeb.Text <> "__/__/____" Then
                    'tempmeb.Text = Format(Now, "dd/mm/yyyy")
                        'Grid1.col = Grid1.col - 1
                           Grid1.col = Grid1.col + 1
                        Grid1.SetFocus
                        Grid1_Click
                    ''''''End If
                    
                    Case 6
                    ''''''''''''If tempmeb.Text <> "__/__/____" Then
                        'Grid1.col = Grid1.col - 1
                        Grid1.col = Grid1.col + 1
                        Grid1.SetFocus
                        Grid1_Click
                    ''''''''End If
                
                    Case 7
                    If Val(tempmeb.Text) > 0 Then
                        'Grid1.col = Grid1.col - 1
                        Grid1.col = Grid1.col + 1
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                    
                    
                    Case 8
                    If tempmeb.Text <> "" Then
                        'Grid1.col = Grid1.col - 1
                        Grid1.col = Grid1.col + 1
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                    
                    
                    Case 9
                    Grid1.col = 1
                    Grid1.row = Grid1.row + 1
                    Grid1.Rows = Grid1.Rows + 1
                    Grid1.SetFocus
                    Grid1_Click
                
                'Case 6
                 '   Grid1.col = 1
                  '  Grid1.row = Grid1.row + 1
                   ' Grid1.Rows = Grid1.Rows + 1
                    'Grid1.SetFocus
                    'Grid1_Click
            End Select
        Else

         'If Grid1.col = 3 Or Grid1.col = 4 Or Grid1.col = 7 Or Grid1.col = 8 Then
                  If Grid1.col = 3 Or Grid1.col = 7 Or Grid1.col = 8 Then
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
    
'''''If Grid1.col = 5 Or Grid1.col = 6 Then
'''''
'''''If Grid1.Text <> "" Then
'''''tempmeb.Text = Grid1.Text
'''''Else
'''''tempmeb.Mask = "__/__/____"
'''''End If
'''''
'''''Else

If Grid1.Text <> "" Then
tempmeb.Mask = ""
tempmeb.Text = ""
tempmeb.Text = Grid1.Text
Else
tempmeb.Mask = ""
tempmeb.Text = ""
End If
'''''''End If

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

Private Sub through1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(Me.customercode.Text) <> "" Then
            Grid1.col = 1
            Grid1.row = 1
            Grid1_Click
        Else
            Me.textbox.SetFocus
            'Me.customercode.SetFocus
        End If
    End If

End Sub

Private Sub through1_LostFocus()
through1 = UCase(through1)
End Sub

'Private Sub Timer1_Timer()
'''''Label20.Move Label20.Left - 20
'''''    If Label20.Left <= -4000 Then
'''''    Label20.Left = 11505
'''''    End If
'End Sub

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
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
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
    
    Open "" + App.Path + "\vipin.txt" For Output As #1
    Line = 0
header:
      If kkk.State = 1 Then
            kkk.Close
      End If
      If flagyes = True Then
      CNSetup
          kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
          If Not kkk.BOF Then
             Print #1, Chr(27) + Chr(18) + Chr(14)
             Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15); Chr(27) + Chr(18) + Chr(14); dspace(Trim(kkk!cname))
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
    rs1.Open "select * from Purchasea where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
   
    Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!subledger; Tab(T5); "Invoice No. : "; Trim(rs1!INVOICENO); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!InvoiceDate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.Close
            End If
            kkk.Open "select * from sledger where subledger='" + Trim(rs1!subledger) + "' and " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE;
                Print #1, Tab(3); kkk!address1; Tab(T5); "Order by    : "; Trim(rs1!ORDERBY); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!OrderDate
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
            kk.Open "select * from Purchaseb where  " & stringyear & "   and invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                        tdata.Open "Select bookname from books where  " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                        tdata.Open "select sum(amount) from Purchaseb where  " & stringyear & "   and invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
                        If Not tdata.BOF Then
                            
                            Print #1, Tab(T7 - 1); rsets(Trim(Format(str(myround(tdata(0), 2)), "0.00")), 12)
                            Print #1, Tab(T5); "Less Discount @ " + Trim(Format(str(myround(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(str(myround(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(str(myround(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                            netamount = netamount + myround(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                        End If
                        tdata.Close
                        Print #1, Tab(T7); repli("-", 22)
                Loop
            End If
           End If
           Print #1, Tab(T5 - 4); rsets(Trim(str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12)
           Print #1, Tab(T6); repli("-", 22)
           If kk.State = 1 Then
                kk.Close
           End If
           kk.Open "Select * from Purchasec where  " & stringyear & "   and invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
           If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DEBITORCREDIT) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5); Trim(kk!Text) + "    " + Trim(Format(str(myround(kk!rate, 2)), "0.00")); Tab(T8 + 5); rsets(Trim(Format(str(myround(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5); Trim(kk!Text); Tab(T8 + 5); rsets(Trim(Format(str(myround(kk!amount, 2)), "0.00")), 12)
                    End If
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T6); repli("-", 22)
            Print #1, Tab(6); Chr(71) + "NET AMOUNT: "; Tab(T8 + 5); Chr(72) + rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12)
           End If
           kk.Close
           kk.Open "Select * from Purchasea where  " & stringyear & "   and invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
           If Not kk.BOF Then
                If kk!txt1a <> 0 Then
                    Print #1, Tab(T5); kk!txt1; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!txt1a, 2))), "0.00")), 12)
                    netamount = netamount + myround(kk!txt1a, 2)
                End If
                If kk!txt2a <> 0 Then
                    Print #1, Tab(T5); kk!txt2; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!txt2a, 2))), "0.00")), 12)
                    netamount = netamount + myround(kk!txt2a, 2)
                End If
                If kk!baa <> 0 Then
                    Print #1, Tab(T5); "BY BANK "; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!baa, 2))), "0.00")), 12)
                    netamount = netamount - myround(kk!baa, 2)
                End If
           End If
           Print #1, Tab(T6); repli("-", 22)
           Print #1, Tab(T5); Chr(71) + "BALANCE : "; Tab(T8 + 5); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12) + Chr(72)
        
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
            tempdata.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
Dim rs As ADODB.Recordset
Dim Pno As Integer
Set rs = New ADODB.Recordset
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
Open "" + App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
FooterYes = False
header:
If kkk.State = 1 Then
      kkk.Close
End If
CNSetup
kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
If printheader = True Then
   If Not kkk.BOF Then
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, Chr(27) + Chr(71); Chr(27) + Chr(15) + Chr(14)
     Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
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
Print #1, Tab(1); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(15) + Chr(14); Tab(30); dspace(Trim("INVOICE")); Chr(20); Tab(T4 + 6); IIf(printheader = True, kkk!uptt, "")
If printheader = True Then
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
rs1.Open "select * from Purchasea where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To, S.L. Code : "; Tab(T1 + 10); Mid$(rs1!subledger, 1, 5); Tab(T5); "Invoice No. : "; Trim(rs1!INVOICENO); Tab(T8); "Dated     : "; rs1!InvoiceDate   'Chr(27) + Chr(18);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from sledger where  " & stringyear & "   and subledger='" + Trim(rs1!subledger) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE
        Print #1, Tab(3); kkk!address1; Tab(T5); "Order by    : "; Trim(rs1!ORDERBY); Tab(T8); "Dated     : "; IIf(IsNull(rs1!OrderDate), "  /  /    ", rs1!OrderDate)
        Print #1, Tab(3); kkk!ADDRESS2; Tab(T5); "Bilty No.   : "; Trim(rs1!biltyno); Tab(T8); "Dated     : "; IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, Tab(3); kkk!ADDRESS3
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
    kk.Open "select * from Purchaseb where  " & stringyear & "   and invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where  " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
                Print #1, rsets(Trim(str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(str(kk!quantity)), 5); Tab(T6 + 1); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
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
                tdata.Open "select sum(amount) from Purchaseb where  " & stringyear & "   and invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                    Print #1, Tab(T7 - 1); rsets(Trim(Format(str(myround(tdata(0), 2)), "0.00")), 12)
                    Print #1, Tab(T5); "Less Discount @ " + Trim(Format(str(myround(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(str(myround(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(str(myround(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                    Print #1, Tab(T7); repli("-", 22)
                    Line = Line + 3
                    netamount = netamount + myround(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                End If
                tdata.Close
                'Print #1, Tab(t7); repli("-", 22)
                'line = line + 1
                Loop
            End If
        End If
        Print #1, repli("-", 145)
        Print #1, Tab(T5 - 4); rsets(Trim(str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12)
       'by vk Print #1, Tab(t8); repli("-", 22)
        Line = Line + 2
        If kk.State = 1 Then
             kk.Close
        End If
        kk.Open "Select * from Purchasec where  " & stringyear & "   and invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DEBITORCREDIT) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5 + 21); Trim(kk!Text) + " :  @  " + Trim(Format(str(myround(kk!rate, 2)), "0.00")) & " % "; Tab(T8 + 5); rsets(Trim(Format(str(myround(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5 + 20); Trim(kk!Text) & " :"; Tab(T8 + 5); rsets(Trim(Format(str(myround(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T8); repli("-", 22)
            Print #1, Chr(27) + Chr(71); Tab(T5 + 20); "NET AMOUNT  : "; Tab(T8 + 6); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
            VNetamt = netamount
        Line = Line + 2
        End If
        kk.Close
        kk.Open "Select * from Purchasea where  " & stringyear & "   and invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(T5 + 20); kk!txt1 & "  :"; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + myround(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(T5 + 20); kk!txt2 & " :"; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - myround(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(T5 + 20); "BY BANK       :"; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - myround(kk!baa, 2)
             End If
        End If
        Print #1, Tab(T8); repli("-", 22)
        Print #1, Tab(T5 + 20); Chr(27) + Chr(71); "BALANCE    : "; Tab(T8 + 6); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
        Print #1, Tab(T8); repli("-", 22)
        Line = Line + 3
       ' PRINT THE FOOTER IN INVOICE START
        Do While Line < 61
            Print #1, ""
            Line = Line + 1
        Loop
        Print #1, Tab(0); toword(myround(VNetamt, 2))

        Print #1, Tab(0); repli("-", 145)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        'Dim LEFTM As Integer
        'LEFTM = 5
        CNSetup
        tempdata.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
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


