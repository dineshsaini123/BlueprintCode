VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form countersale 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7095
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "cashsale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbAgentName 
      Height          =   315
      Left            =   6810
      TabIndex        =   25
      Top             =   4950
      Width           =   2745
   End
   Begin VB.ComboBox cmbdiscountcat 
      Height          =   315
      Left            =   6540
      TabIndex        =   6
      Top             =   390
      Width           =   1275
   End
   Begin VB.ComboBox Combosldistrictcode 
      Height          =   315
      Left            =   5400
      TabIndex        =   10
      Top             =   720
      Width           =   1395
   End
   Begin VB.Frame frame1 
      Height          =   675
      Left            =   60
      TabIndex        =   62
      Top             =   -60
      Width           =   1785
      Begin VB.OptionButton Optioncredit 
         Caption         =   "Credit"
         Height          =   255
         Left            =   930
         TabIndex        =   2
         Top             =   270
         Width           =   705
      End
      Begin VB.OptionButton Optioncash 
         Caption         =   "Cash"
         Height          =   225
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin MSMask.MaskEdBox textbox 
      Height          =   315
      Left            =   6180
      TabIndex        =   5
      Top             =   30
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Commandall 
      Caption         =   "All Books"
      Height          =   375
      Left            =   1050
      TabIndex        =   34
      Top             =   5310
      Width           =   945
   End
   Begin VB.CommandButton Commandother 
      Caption         =   "&Other"
      Height          =   375
      Left            =   210
      TabIndex        =   33
      Top             =   5310
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2685
      Left            =   90
      TabIndex        =   17
      Top             =   1050
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   4736
      _Version        =   393216
      FillStyle       =   1
      Appearance      =   0
   End
   Begin VB.ComboBox Bookname 
      Height          =   960
      Left            =   3480
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   13
      Top             =   2550
      Width           =   2295
   End
   Begin VB.ComboBox Bookcode 
      Height          =   765
      Left            =   420
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   12
      Top             =   2550
      Width           =   2355
   End
   Begin VB.PictureBox Picture5 
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   9495
      TabIndex        =   15
      Top             =   5760
      Width           =   9555
      Begin VB.CommandButton Commandprintnh 
         Caption         =   "N&HPrint"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6180
         TabIndex        =   46
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Commandadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1050
         TabIndex        =   0
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   375
         Left            =   8130
         TabIndex        =   50
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7140
         TabIndex        =   48
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5310
         TabIndex        =   45
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   43
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   375
         Left            =   3600
         TabIndex        =   41
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Height          =   375
         Left            =   2760
         TabIndex        =   39
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1890
         TabIndex        =   37
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   210
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   800
      End
   End
   Begin MSMask.MaskEdBox through 
      Height          =   315
      Left            =   3660
      TabIndex        =   24
      Top             =   4950
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox I_DTOB 
      Height          =   315
      Left            =   3210
      TabIndex        =   9
      Top             =   690
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox weight 
      Height          =   315
      Left            =   1410
      TabIndex        =   22
      Top             =   4950
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox freight 
      Height          =   315
      Left            =   30
      TabIndex        =   21
      Top             =   4950
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bdated 
      Height          =   315
      Left            =   5580
      TabIndex        =   20
      Top             =   4290
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox biltno 
      Height          =   315
      Left            =   3660
      TabIndex        =   19
      Top             =   4290
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bundles 
      Height          =   285
      Left            =   7980
      TabIndex        =   11
      Top             =   750
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox I_OB 
      Height          =   285
      Left            =   1110
      TabIndex        =   8
      Top             =   690
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
      Left            =   3330
      TabIndex        =   4
      Top             =   330
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox station 
      Height          =   315
      Left            =   30
      TabIndex        =   18
      Top             =   4290
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tempmeb 
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Top             =   2040
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
      TabIndex        =   16
      Top             =   3180
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
      TabIndex        =   52
      Top             =   2640
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
      Left            =   3330
      TabIndex        =   3
      Top             =   30
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox through1 
      Height          =   285
      Left            =   4770
      TabIndex        =   26
      Top             =   4950
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.ComboBox Genledger 
      Height          =   315
      Left            =   9540
      Sorted          =   -1  'True
      TabIndex        =   53
      Top             =   1110
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSMask.MaskEdBox marka 
      Height          =   315
      Left            =   2460
      TabIndex        =   23
      Top             =   4950
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin VB.ComboBox customercode 
      Height          =   960
      Left            =   6150
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agent :"
      Height          =   315
      Left            =   6810
      TabIndex        =   66
      Top             =   4590
      Width           =   2760
   End
   Begin VB.Label lbldis 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Discount Category"
      Height          =   285
      Left            =   4740
      TabIndex        =   65
      Top             =   390
      Width           =   1695
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Railway/Station : "
      Height          =   285
      Left            =   30
      TabIndex        =   64
      Top             =   4020
      Width           =   3600
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "District Name"
      Height          =   285
      Left            =   4320
      TabIndex        =   63
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label labelbybanklbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By Cash : "
      Height          =   255
      Left            =   2400
      TabIndex        =   61
      Top             =   5310
      Width           =   1200
   End
   Begin VB.Label labelbybank 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3630
      TabIndex        =   60
      Top             =   5310
      Width           =   1200
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   5580
      TabIndex        =   29
      Top             =   4050
      Width           =   1200
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marka : "
      Height          =   285
      Left            =   2460
      TabIndex        =   36
      Top             =   4620
      Width           =   1155
   End
   Begin VB.Label mgd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8070
      TabIndex        =   56
      Top             =   3750
      Width           =   1230
   End
   Begin VB.Label mna 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8100
      TabIndex        =   55
      Top             =   4050
      Width           =   1200
   End
   Begin VB.Label mga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6840
      TabIndex        =   54
      Top             =   3750
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   1920
      TabIndex        =   51
      Top             =   330
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cash Memo No. : "
      Height          =   285
      Left            =   1920
      TabIndex        =   49
      Top             =   30
      Width           =   1335
   End
   Begin VB.Label label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cust Code : "
      Height          =   285
      Left            =   4740
      TabIndex        =   47
      Top             =   60
      Width           =   1425
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Net Amount : "
      Height          =   255
      Left            =   6810
      TabIndex        =   44
      Top             =   4050
      Width           =   1200
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Gross Amount : "
      Height          =   255
      Left            =   6660
      TabIndex        =   42
      Top             =   3390
      Width           =   1200
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Order By : "
      Height          =   285
      Left            =   90
      TabIndex        =   40
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   2130
      TabIndex        =   38
      Top             =   690
      Width           =   1050
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bundle(s):"
      Height          =   285
      Left            =   6840
      TabIndex        =   32
      Top             =   750
      Width           =   1065
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Through : "
      Height          =   285
      Left            =   3660
      TabIndex        =   31
      Top             =   4620
      Width           =   3135
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bilty No. : "
      Height          =   285
      Left            =   3660
      TabIndex        =   30
      Top             =   4050
      Width           =   1875
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Freight : "
      Height          =   285
      Left            =   0
      TabIndex        =   28
      Top             =   4620
      Width           =   1395
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Weight : "
      Height          =   285
      Left            =   1410
      TabIndex        =   27
      Top             =   4620
      Width           =   1050
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Discount : "
      Height          =   255
      Left            =   7890
      TabIndex        =   57
      Top             =   2880
      Width           =   1290
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Quantity : "
      Height          =   255
      Left            =   2160
      TabIndex        =   59
      Top             =   3750
      Width           =   1470
   End
   Begin VB.Label tqu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3660
      TabIndex        =   58
      Top             =   3750
      Width           =   840
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
Dim FooterYes As Boolean
Dim totalquantity As Long
Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
Dim rs As ADODB.Recordset
Dim LEFTM As Integer
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
paperWidth = 81
MaxLine = 60
called1 = False
called2 = False
Dim Line As Integer
Dim rs1 As ADODB.Recordset
Dim kkk As ADODB.Recordset
Set kkk = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Open "" + App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
LEFTM = 5
FooterYes = False
header:
    If kkk.State = 1 Then
          kkk.Close
    End If
    CNSetup
    kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If FooterYes = True Then
        If Line > MaxLine - 5 Then
            Do While Line < 61
                Print #1, ""
                Line = Line + 1
            Loop
        End If
        Line = 0
        LEFTM = 5
        Print #1, Tab(0); repli("-", 81)
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
        
End If
If printheader = True Then
   If Not kkk.BOF Then
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, Chr(27) + Chr(71); Chr(27) + Chr(77) + Chr(14)
     Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 17); Chr(27) + Chr(77) + Chr(14); dspace(Trim(kkk!cname))
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
Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CASH MEMO")))) / 2 - 8); Chr(14); "***CASH MEMO***"; Chr(20)
Line = Line + 1
If printheader = True Then
   Print #1, Tab(48); kkk!uptt
   Print #1, Tab(48); kkk!cst
   Line = Line + 2
End If
If printheader = False Then
   Print #1, ""
   Print #1, ""
   Line = Line + 2
End If
Print #1, repli("-", 81)
Line = Line + 1

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from casha where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"

If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); " To,"; Tab(7); IIf(Optioncash.Value = True, "", Mid$(rs1!subledger, 1, 5)); Tab(38); "Cash Memo No.: "; Trim(rs1!INVOICENO); Tab(67); "Dt. : "; rs1!InvoiceDate; Chr(27) + Chr(72);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from sledger where   " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & IIf(Optioncash.Value = True, rs1!CASHPARTYNAME, kkk!DESCFORINVOICE)
        Print #1, Tab(5); IIf(IsNull(kkk!address1), "", kkk!address1); Tab(37); Chr(27) + Chr(71); "Order by     : "; Chr(27) + Chr(72); Trim(rs1!ORDERBY); Tab(68); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!OrderDate), "  /  /    ", rs1!OrderDate)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS2), "", kkk!ADDRESS2); Tab(37); Chr(27) + Chr(71); "Bilty No.    : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(68); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS3), "", kkk!ADDRESS3)
        kkk.Close
        Print #1, Chr(27) + Chr(71); "Through  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!through) + IIf(Trim(rs1!through1) = "", "", "," & rs1!through1)
        Print #1, Chr(27) + Chr(71); "Station  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!station); Tab(56); Chr(27) + Chr(71); "Pvt. Mark   : "; Chr(27) + Chr(72); Trim(rs1!marka)
        Print #1, Chr(27) + Chr(71); "Freight  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!freight); Tab(35); Chr(27) + Chr(71); "Weight    : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."; Tab(58); Chr(27) + Chr(71); "Bundle(s)   : "; Chr(27) + Chr(72); Trim(rs1!bundles)
        Print #1, Chr(27) + Chr(71); repli("-", 81)
        Print #1, Tab(0); "S.No."; Tab(10); "Book Description"; Tab(44); "Qty."; Tab(52); "Rate"; Tab(61); "Amount"; Tab(71); "Net Amount"
        Print #1, repli("-", 81); Chr(27) + Chr(72)
        Line = Line + 10
    End If
    If called1 Then
        GoTo printagain1
    End If
    If called2 Then
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.Close
    kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,sno ", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                'Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(Str(kk!quantity)), 5); Tab(T6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                Print #1, Tab(0); rsets(Trim(str(sno)), 4); Tab(6); Trim(tdata!Bookname); Tab(41); rsets(Trim(str(kk!quantity)), 5); Tab(48); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(56); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
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
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
                    
                    
printagain2:
                    called2 = False
                End If
                Print #1, Tab(57); repli("-", 12)
                Line = Line + 1
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CashB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(56); rsets(Trim(Format(str(tdata(0)), "0.00")), 12)
                   'Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(Str(tdata(0) * cdiscount / 100), "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata(0) - (tdata(0) * cdiscount / 100)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(str(cdiscount), "0.00")) + " %"; Tab(56); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(69); rsets(Trim(Format(str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(57); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.Close
             Loop
         End If
    End If
    Print #1, repli("-", 81)
    Print #1, Tab(41); rsets(Trim(str(totalquantity)), 5); Tab(69); rsets(Trim(Format(str(netamount), "0.00")), 12)
    Line = Line + 2
    If kk.State = 1 Then kk.Close
    kk.Open "Select * from CASHC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kk.BOF Then
       Do While Not kk.EOF
             If kk!amount > 0 Then
                    If Trim(kk!DEBITORCREDIT) = Trim("Credit") Then
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
        kk.Close
        kk.Open "Select * from CASHA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                 Print #1, Tab(48); Chr(27) + Chr(71); "BALANCE     : "; Tab(70); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
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
        Print #1, Tab(0); Chr(27) + Chr(71); toword(myround(VNetamt, 2)); Chr(27) + Chr(72)
        Print #1, Tab(0); repli("-", 81)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        CNSetup
        tempdata.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
       ' Print #1, ""
        'PRINT THE FOOTER IN INVOICE END
        Close #1
        PrintOption.Show
        
End Sub

Sub invoicecalc()
'OTHERCASH.calc
     mga.Caption = Format(myround(totalamount, 2), "0.00")
    ' mgd.Caption = Format(myround(totaldiscount, 2), "0.00")
     
     mna.Caption = Format(myround((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
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
  '  kk.Open "SELECT MAX(INVOICENO) FROM casha where " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
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
                If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) Then
                    ctl.Text = ""
                End If
                ctl.Enabled = False
            End If
        Next
        For i = 1 To maxrow
           Grid1.row = i
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
                rs.Open "select * from books where   " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
                row = Grid1.row
                col = Grid1.col
                If Trim(Grid1.Text) <> "" Then
                    If Not rs.BOF Then
                        rs.MoveFirst
                        rs.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If rs.EOF Then
                            tempmeb.Visible = True
                            tempmeb.SetFocus
                            rs.Close
                            templost = False
                            Exit Function
                        Else
                            Grid1.Text = rs(0)
                            Grid1.col = 2
                            Grid1.Text = rs(1)
                         '   If Not edit Then
                                Grid1.col = 3
                                If Trim(Grid1.Text) = "" Then
                                    Grid1.Text = 0
                                End If
                                q = Val(Grid1.Text)
                                Grid1.col = 5
                                If Trim(Grid1.Text) = "" Then
                                Grid1.Text = Format(rs(3), "0.00")            'rs(3)
                                r = rs(3)
                                End If
                                '/******************
                                  Set kk = CON.Execute("select DISCATEGORY from sledger where   " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
                                
                                Grid1.col = 6
                                If Grid1.Text = "" Or addmode = True Then
                                If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    kk.Close
                                   If Optioncash.Value = 0 Then
                                    Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
                                    
                                    Else
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(rs(2)) + "'")
                                    End If
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
                            End If
                            Grid1.col = col
                            rs.Close
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
                Grid1.Text = Format(myround(q * r, 2), "0.00")
                Grid1.col = 8
                Grid1.Text = Format(myround((q * r) * (D / 100), 2), "0.00")
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
        For i = 1 To maxrow
            Grid1.row = i
            Grid1.col = 7
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
        Next
        invoicecalc
        Me.tqu.Caption = ""
        For i = 1 To maxrow
            Grid1.col = 3
            Grid1.row = i
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
                           'Commandother.SetFocus
                           station.SetFocus
                           
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
                rs.Open "select * from books where   " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
                row = Grid1.row
                col = Grid1.col
                If Trim(Grid1.Text) <> "" Then
                    If Not rs.BOF Then
                        rs.MoveFirst
                        rs.Find "bookname='" + Trim(Grid1.Text) + "'"
                        If rs.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            rs.Close
                            Exit Sub
                        Else
                            
                            Grid1.col = 1
                            Grid1.Text = rs(0)
                            Grid1.col = 2
                            Grid1.Text = rs(1)
                        '    If Not edit Then
                                 Grid1.col = 3
                                If Trim(Grid1.Text) = "" Then
                                        Grid1.Text = 0
                                End If
                                q = Val(Grid1.Text)
                                Grid1.col = 5
                                Grid1.Text = Format(rs(3), "0.00")
                                r = rs(3)
                                '/******************
                                Set kk = CON.Execute("select DISCATEGORY from sledger where   " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
                                
                                Grid1.col = 6
                                If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    kk.Close
                                    If Optioncash.Value = 0 Then
                                    Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
                                    
                                    Else
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(rs(2)) + "'")
                                    End If
                                    
                                    
                                    
                                 '   Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
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
                                Grid1.Text = myround(q * r, 2)
                                Grid1.col = 8
                                Grid1.Text = myround((q * r) * (D / 100), 2)
                         '   End If
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
  rs1.Open "select * from agentmaster where " & stringyear & " and agentname='" & cmbAgentName.Text & "' order by agentname", CON, adOpenKeyset, adLockReadOnly, adCmdText
  If rs1.RecordCount <= 0 Then
     MsgBox "Enter Valid Agent Name.. "
     cmbAgentName.SetFocus
  End If
End If

End Sub

Private Sub Combosldistrictcode_LostFocus()
If Combosldistrictcode.Text = "" Then
   Combosldistrictcode.SetFocus
   Exit Sub
End If
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
If Combosldistrictcode.Text <> "" And addmode = True Then
   rs1.Open "Select * from Districts where  " & stringyear & " and Districtname = '" & Combosldistrictcode.Text & "'", CON, adOpenStatic, adLockReadOnly
   If rs1.RecordCount > 0 Then
      Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
   End If
End If


End Sub

Private Sub Commandprintnh_Click()
printheader = False
printinvoice
End Sub

Private Sub Commandabandon_Click()
invoiceabandon
Me.Commandall.Enabled = False
Me.Commandother.Enabled = False
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub
Private Sub Commandadd_Click()
    invoiceabandon
    
    
    
    Dim rs As ADODB.Recordset
    addoredit = True
    addmode = True
    edit = False
    
    Set rs = New ADODB.Recordset
    Dim TEMPNUM As Integer
    If edit = False Then
    If CON.Execute("SELECT MAX(INVOICENO) FROM casha where " & stringyear)(0) >= Val(Trim(Me.I_NO.Text)) Then
        Me.I_NO.Text = CON.Execute("SELECT MAX(INVOICENO) FROM casha where " & stringyear)(0) + 1
         rs.Open "tempCash", CON, adOpenKeyset, adLockOptimistic, adCmdText
         If rs.BOF Then
             rs.addNew
         End If
         rs!In = Val(Me.I_NO.Text)
         rs.Update
         rs.Close
        
    End If
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
    Commandsave.Enabled = True
    Commandsearch.Enabled = False
    Grid1.Enabled = True
    Me.customercode.Enabled = True
    Me.Optioncash.SetFocus
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
    rs.Open "select * from books where   " & stringyear & " order by bookcode", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
            Set kk = CON.Execute("select DISCATEGORY from sledger where   " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
            Grid1.col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.Close
                Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
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

End Sub

Private Sub Commanddelete_Click()
If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
                CON.Execute ("DELETE from CASHA where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("DELETE from CASHB where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("DELETE from CASHC where INVOICENO = " + Trim(I_NO.Text))
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
    Commandsave.Enabled = True
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
    CON.Execute ("DELETE from CASHCTMP")
    DoEvents
    CON.Execute ("insert into CASHCTMP  select * from CASHC where INVOICENO = " + Trim(I_NO.Text))
    'invoicetmp creation end
    addoredit = False
    HIT
    VB.Screen.ActiveForm.ActiveControl.SelLength = 2
    OTHERCASH.TOP = 0
    OTHERCASH.Left = 0
    OTHERCASH.Visible = False
    
End Sub
Private Sub Commandother_Click()
'    Me.Enabled = False
    OTHERCASH.TOP = 0
    OTHERCASH.Left = 0
    'Unload OTHERCASH
    'Load OTHERCASH
    OTHERCASH.Show
End Sub
Private Sub CommandPrint_Click()
   printheader = True
   printinvoice
End Sub
Private Sub Commandreturn_Click()
    Unload Me
    addoredit = False
    MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()
If Optioncash = True Then
    If Trim(Combosldistrictcode) = "" Then
        MsgBox "Please Enter District"
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
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
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
If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(customercode.Text) <> "" Then
   If edit Then
      CON.Execute ("DELETE from CASHA where INVOICENO = " + Trim(I_NO.Text))
      CON.Execute ("DELETE from CASHB where INVOICENO = " + Trim(I_NO.Text))
      CON.Execute ("DELETE from CASHC where INVOICENO = " + Trim(I_NO.Text))
   End If
   If rs.State = 1 Then rs.Close
   LAMOUNT = 0
   rs.Open "select * from casha where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
   If Not edit Then
again:
      If CON.Execute("SELECT MAX(INVOICENO) FROM casha where " & stringyear)(0) >= Val(Trim(Me.I_NO.Text)) Then
        ' Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
         'GoTo again
      End If
   End If
   rs.addNew
   rs!INVOICENO = Val(Me.I_NO.Text)
   rs!InvoiceDate = Me.i_dt.Text
   rs!Genledger = Trim(Me.Genledger.Text)
   rs!subledger = Trim(Me.customercode.Text)
   rs!ORDERBY = Trim(Me.I_OB.Text)
   If Trim(Me.I_DTOB) = Trim("__/__/____") Then
     
      rs!OrderDate = Null
   Else
      rs!OrderDate = Trim(Me.I_DTOB.Text)
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
   If Trim(Me.bdated) = Trim("__/__/____") Then
      rs!BILTYDATE = Null
      'rs!BILTYDATE = Date
   Else
     rs!BILTYDATE = Me.bdated & ""
   End If
   rs!freight = Me.freight & ""
   rs!weight = Me.weight & ""
   rs!netamount = myround(Val(Trim(Me.mna.Caption)), 2)
   rs!gamount = (Me.totalamount - Me.totaldiscount)
   rs!txt1 = Trim(OTHERCASH.T1TEXT.Text)
   rs!txt1a = Val(Trim(OTHERCASH.T1.Text))
   rs!txt2 = Trim(OTHERCASH.T2TEXT.Text)
   rs!txt2a = Val(Trim(OTHERCASH.T2.Text))
   'rs!baa = Val(Trim(OTHERCASH.T3TEXT.Text))
   
   rs!baa = Val(Trim(OTHERCASH.T3TEXT.Text))
   rs!baa = Val(Trim(labelbybank.Caption))
   
   rs!District = Combosldistrictcode.Text
   rs!CASHPARTYNAME = textbox.Text
   rs!agentname = cmbAgentName.Text
   rs!discat = cmbdiscountcat.Text
  
err1:
   If Not edit Then
      If CON.Execute("SELECT MAX(INVOICENO) FROM casha where " & stringyear)(0) >= Val(Trim(Me.I_NO.Text)) Then
         'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
         'rs!INVOICENO = Val(Me.I_NO.Text)
         On Error GoTo err1
      End If
   End If
   rs.Update
   On Error GoTo 0
   rs.Close
   rs.Open "select * from cashb where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
   Dim i As Integer
   RRRR = Grid1.row
   CCCC = Grid1.col
   For i = 1 To maxrow
       Grid1.row = i
       Grid1.col = 1
       If Trim(Grid1.Text) <> "" Then
          Grid1.col = 3
          If Val(Trim(Grid1.Text)) > 0 Then
             Grid1.col = 5
            If Val(Trim(Grid1.Text)) > 0 Then
               rs.addNew
               Grid1.col = 1
               rs!INVOICENO = Val(Me.I_NO.Text)
               rs!InvoiceDate = Me.i_dt.Text
               rs!Genledger = Trim(Me.Genledger.Text)
               rs!subledger = Trim(Me.customercode.Text)
               rs!Bookcode = Trim(Grid1.Text)
               Grid1.col = 3
               rs!quantity = Trim(Grid1.Text)
               Grid1.col = 5
               rs!rate = Trim(Grid1.Text)
               Grid1.col = 7
               rs!amount = Trim(Grid1.Text)
               LAMOUNT = Val(Trim(Grid1.Text))
               Grid1.col = 4
               rs!PrintOrder = Trim(Grid1.Text)
               Grid1.col = 6
               rs!discount = Trim(Grid1.Text)
               Grid1.col = 8
               rs!netamount = LAMOUNT - Trim(Grid1.Text)
               LAMOUNT = 0
               rs.Update
            End If
         End If
     End If
  Next
  rs.Close
  Grid1.TopRow = 1
  rs.Open "select * from cashc where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
  '/******
  'Dim I, x As Integer
   Dim temprs As ADODB.Recordset
   Set temprs = New ADODB.Recordset
       For i = 1 To OTHERCASH.mrow
           OTHERCASH.Grid1.row = i
           OTHERCASH.Grid1.col = 0
           If Trim(OTHERCASH.Grid1.Text) <> "" Then
              rs.addNew
              rs!INVOICENO = Val(Me.I_NO.Text)
              rs!InvoiceDate = Me.i_dt.Text
              rs!gamount = (Me.totalamount - Me.totaldiscount)
              rs!Text = Trim(OTHERCASH.Grid1.Text)
              If temprs.State = 1 Then temprs.Close
              If edit Then
                 temprs.Open "select * from CASHCTMP", CON, adOpenKeyset, adLockReadOnly, adCmdText
                 If OTHERCASH.Grid1.Text <> "" Then
                    temprs.Find "TEXT='" + Trim(OTHERCASH.Grid1.Text) + "'"
                    rs!Genledger = temprs!Genledger & ""
                    rs!subledger = temprs!subledger & ""
                    rs!DEBITORCREDIT = temprs!DEBITORCREDIT & ""
                    rs!RYN = temprs!RYN & ""
                End If
                temprs.Close
              Else
                 temprs.Open "select * from cashend where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
                 If OTHERCASH.Grid1.Text <> "" Then
                    temprs.Find "TEXT='" + Trim(OTHERCASH.Grid1.Text) + "'"
                    rs!Genledger = temprs!Genledger & ""
                    rs!subledger = temprs!subledger & ""
                    rs!DEBITORCREDIT = temprs!DEBITORCREDIT & ""
                    rs!RYN = temprs!RYN & ""
                 End If
                 temprs.Close
              End If
              OTHERCASH.Grid1.col = 1
              rs!rate = Val(Trim(OTHERCASH.Grid1.Text))
              If Val(Trim(OTHERCASH.Grid1.Text)) > 0 Then
                 rs!amount = (Me.totalamount - Me.totaldiscount) * (Val(Trim(OTHERCASH.Grid1.Text)) / 100)
              Else
                OTHERCASH.Grid1.col = 2
                rs!amount = Val(Trim(OTHERCASH.Grid1.Text))
              End If
              rs.Update
          End If
      Next
      rs.Close
      If addmode = True Then
         rs.Open "tempCash", CON, adOpenKeyset, adLockOptimistic, adCmdText
         If rs.BOF Then
             rs.addNew
         End If
         rs!In = Val(Me.I_NO.Text)
         rs.Update
         rs.Close
      End If
         SAVED = True
  End If
  If SAVED Then
      Unload OTHERCASH
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
    
 End Sub

Private Sub Commandsearch_Click()
   Unload OTHERCASH
    Me.Enabled = False
    'searchscreen.Grid1.row = 0
    'searchscreen.Grid1.col = 0
    Call searchscreen.tempr(17, "countersale")
End Sub

Private Sub customercode_KeyPress(KeyAscii As Integer)
   ' If KeyAscii = 13 Then
  '  SendKeys "{DOWN}"
   ' SendKeys "{TAB}"
   ' marka.SetFocus
   ' End If
End Sub
Private Sub customercode_LostFocus()
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open "select * from sledger where   " & stringyear & " and gledger='SUNDRY DEBTORS' and subledger='" + Trim(customercode.Text) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.RecordCount <= 0 Then
           customercode.SetFocus
           HIT
           rs.Close
           Exit Sub
        End If
        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        If rs!distcode <> "" And addmode = True Then
            rs1.Open "Select * from Districts where  " & stringyear & " and Districtname = '" & rs!distcode & "'", CON, adOpenStatic, adLockReadOnly
            If rs1.RecordCount > 0 Then
                Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
            End If
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
    mna.Enabled = True
    Label2.Enabled = True
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If Grid1.row >= 1 Then
           Grid1.RemoveItem Grid1.row - 1
           a = Grid1.Text
           'tempmeb.Text = a
           a = templost
           Grid1.SetFocus
          End If
   End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If UCase(Trim(VB.Screen.ActiveControl.Name)) = UCase(Trim("CUSTOMERCODE")) Then
         '   SendKeys "{DOWN}"
         
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
    Grid1.TOP = 1200
   ' Set CON = New ADODB.Connection
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Me.TOP = 0
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
    rs.Open "select * from books where  " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.Bookcode.AddItem rs(0)
            Me.Bookname.AddItem IIf(IsNull(rs(1)), "", rs(1))
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "select Distinct categorycode from DISCCATS order by categorycode", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.cmbdiscountcat.AddItem rs!categorycode
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    Genledger.Text = "SUNDRY DEBTORS"
    rs.Open "select * from sledger where   " & stringyear & " and gledger='" + Trim(Genledger.Text) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
    rs.Open "select  Agentname from AgentMaster where " & stringyear & "  order by agentname", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
    rs.Open "SELECT * FROM tempcash WHERE " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rs.BOF Then
       Me.I_NO.Text = rs!In
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
       kk.Open "SELECT MAX(INVOICENO) FROM casha where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
       If kk(0) <> "" Then
          Me.I_NO.Text = Trim(str(kk(0) + 1))
       Else
          Me.I_NO.Text = "1"
       End If
       kk.Close
   End If
   rs.Close
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
    If rs.State = 1 Then rs.Close
    rs.Open "select * from DISTRICTS  where " & stringyear & " order by DISTRICTNAME", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.Combosldistrictcode.AddItem rs!DISTRICTNAME
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub

Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub

 Sub Grid1_Click()
If Trim(Me.customercode.Text) <> "" Then
Dim PREVROW As Integer
Dim prevcol As Integer
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
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
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
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
                tempmeb.TOP = Grid1.TOP + Grid1.CellTop '- 50
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.Text = Grid1.Text
                Bookname.TOP = Grid1.TOP + Grid1.CellTop
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
   PopupMenu dd, , Grid1.Left + X, Grid1.TOP + Y
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
   SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If Trim(I_NO.Text) = "" Then
        MsgBox "Cash Memo No cannot be null"
        I_NO.SetFocus
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "Select * from  CASHA where INVOICENO = " + Trim(I_NO.Text) + "", CON, adOpenStatic, adLockReadOnly
        If rs.EOF Then
            If addoredit = False Then
                 MsgBox "Cash Memo No not found"
                 Exit Sub
            End If
            Exit Sub
        End If
        If addoredit Then
            MsgBox "Cash Memo No already exist..."
            I_NO.SetFocus
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
        I_NO.Text = rs!INVOICENO
        Me.i_dt.Text = rs!InvoiceDate
        Me.Genledger.Text = Trim(rs!Genledger)
        Me.customercode.Text = Trim(rs!subledger)
        Me.textbox.Text = Trim(rs!subledger)
        Me.I_OB.Text = Trim(rs!ORDERBY)
        Me.I_DTOB.Text = IIf(IsNull(rs!OrderDate), "__/__/____", rs!OrderDate)
        Me.marka.Text = Trim(rs!marka)
        Me.bundles = Trim(rs!bundles)
        Me.through.Text = rs!through
        Me.through1.Text = rs!through1
        Me.station.Text = rs!station
        Me.biltno.Text = Trim(rs!biltyno)
        Me.bdated = IIf(IsNull(rs!BILTYDATE), "__/__/____", rs!BILTYDATE)
        Me.freight = Trim(rs!freight)
        Me.weight = Trim(rs!weight)
        Me.labelbybank = Format(myround(Val(rs!baa), 2), "0.00")
        mna.Caption = Format(myround(Val(rs!netamount), 2), "0.00")
      
      
        
        If rs!District <> "" Then
            Combosldistrictcode.Text = rs!District
        End If
        textbox.Text = rs!CASHPARTYNAME
        If Me.customercode.Text = "CASH PARTY" Then
            Optioncash = True
            Me.cmbAgentName.Text = IIf(IsNull(rs!agentname), "", Trim(rs!agentname))
        Else
            Optioncredit = True
            Me.cmbAgentName.Text = IIf(IsNull(rs!agentname), "", Trim(rs!agentname))
        End If
        cmbdiscountcat.Text = IIf(IsNull(rs!discat), "", rs!discat)
        rs.Close
        Grid1.TopRow = 1
    '*/**/*/*/*/*//*/*
    If rs.State = 1 Then rs.Close
    rs.Open "Select * from CASHB where INVOICENO =" + Trim(I_NO.Text) + "", CON, adOpenStatic, adLockReadOnly
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
                kk.Open "select * from books where  " & stringyear & " and bookcode='" + Trim(rs!Bookcode) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
                Grid1.col = 2
                Grid1.Text = Trim(kk!Bookname)
                Grid1.col = 3
                Grid1.Text = Trim(rs!quantity)
                Grid1.col = 5
                Grid1.Text = Format(myround(rs!rate, 2), "0.00")
                Grid1.col = 7
                Grid1.Text = Format(myround(rs!amount, 2), "0.00")
                Grid1.col = 4
                Grid1.Text = Format(myround(rs!PrintOrder, 2), "0.00")
                Grid1.col = 6
                Grid1.Text = Format(myround(rs!discount, 2), "0.00")
                Grid1.col = 8
                Grid1.Text = Format(myround(rs!amount * (rs!discount / 100), 2), "0.00")
                End If
                If Not rs.EOF Then
                    rs.MoveNext
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
        For i = 1 To maxrow
            Grid1.row = i
            Grid1.col = 7
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
        Next
        mga.Caption = Format(myround(totalamount, 2), "0.00")
        mgd.Caption = Format(myround(totaldiscount, 2), "0.00")
        Me.tqu.Caption = ""
        For i = 1 To maxrow
            Grid1.col = 3
            Grid1.row = i
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
trs.Open " SELECT DISTCODE    FROM SLEDGER  WHERE   " & stringyear & " and SUBLEDGER='" & customercode.Text & "'", CON, adOpenStatic, adLockOptimistic, adCmdText
       If Not trs.BOF Then
           If Combosldistrictcode.Text = "" Then
               Combosldistrictcode.Text = trs!distcode
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
If Optioncash.Value = True Then
       Label4.Visible = True
       Combosldistrictcode.Visible = True
       Label4.Left = 4220
       Combosldistrictcode.Left = 5350
       Label11.Left = 6950
       bundles.Left = 8040
       cmbdiscountcat.Visible = True
       lbldis.Visible = True
       'cmbAgentName.Visible = False
       'Label13.Visible = False
 End If


End Sub

Private Sub Optioncredit_Click()
If Optioncredit.Value = True Then
       Label4.Visible = False
       Combosldistrictcode.Visible = False
       Label11.Left = 4220
       bundles.Left = 5350
       lbldis.Visible = False
       cmbdiscountcat.Visible = False
       'cmbAgentName.Visible = True
       'Label13.Visible = True

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
        Dim rs As ADODB.Recordset
           Set rs = New ADODB.Recordset
            Select Case Grid1.col
                Case 1
                    If rs.State = 1 Then
                        rs.Close
                    End If
                    rs.Open "select * from books where  " & stringyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
    MaxLine = 50
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
             Print #1, Chr(27) + Chr(15) + Chr(14)
             Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
             Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
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
       Print #1, Chr(27) + Chr(15)
       Line = Line + 4
  End If
  
  
  
    If rs1.State = 1 Then
        rs1.Close
    End If
    rs1.Open "select * from casha where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
   
    Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!subledger; Tab(T5); "Cash Memo No. : "; Trim(rs1!INVOICENO); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!InvoiceDate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.Close
            End If
            kkk.Open "select * from sledger where   " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE
                Print #1, Tab(3); kkk!address1; Tab(T5); "Order by : "; Trim(rs1!ORDERBY); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!OrderDate
                Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.: "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!BILTYDATE
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
            kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                        tdata.Open "Select bookname from books where " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                        tdata.Open "select sum(amount) from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
           kk.Open "Select * from CASHC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
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
           kk.Open "Select * from CASHA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
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




'****************************56


Sub bakupprintinvoice()
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
MaxLine = 66
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
If printheader = True Then
CNSetup
kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
If Not kkk.BOF Then
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, Chr(27) + Chr(71); Chr(27) + Chr(15) + Chr(14)
      Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
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
If printheader = True Then
   Print #1, Tab(T7 + 4); IIf(printheader = True, kkk!uptt, "")
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
   rs1.Close
End If

If rs1.State = 1 Then
    rs1.Close
End If
rs1.Open "select * from casha where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To,"; Tab(T1 - 3); IIf(Optioncash.Value = True, rs1!CASHPARTYNAME, rs1!subledger); Tab(T5); "Cash Memo No. : "; Trim(rs1!INVOICENO); Tab(T8 + 5); "Dt. :"; Tab(T8 + 12); rs1!InvoiceDate 'Chr(27) + Chr(15);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from sledger where   " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); kkk!DESCFORINVOICE
        Print #1, Tab(3); kkk!address1; Tab(T5); "Order by     : "; Trim(rs1!ORDERBY); Tab(T8 + 5); "Dt. :"; Tab(T8 + 12); rs1!OrderDate
        Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.    : "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. :"; Tab(T8 + 12); rs1!BILTYDATE
        kkk.Close
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
        kk.Close
    End If
    kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                  Line = Line + 1
                    GoTo header
printagain2:
                    called2 = False
                End If
                Print #1, Tab(T7); repli("-", 22)
                Line = Line + 1
                tdata.Open "select sum(amount) from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                    Print #1, Tab(T7 - 1); rsets(Trim(Format(str(myround(tdata(0), 2)), "0.00")), 12)
                    Print #1, Tab(T5); "Less Discount @ " + Trim(Format(str(myround(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(str(myround(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(str(myround(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                    Line = Line + 2
                    netamount = netamount + myround(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                End If
                tdata.Close
                Print #1, Tab(T7); repli("-", 22)
                Line = Line + 1
                Loop
            End If
        End If
        Print #1, Tab(T5 - 10); repli("-", 22)
        Print #1, Tab(T5 - 4); rsets(Trim(str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12)
        
        Line = Line + 2
        
        
        If kk.State = 1 Then
             kk.Close
        End If
        kk.Open "Select * from CASHC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T8); repli("-", 22)
            Print #1, Chr(27) + Chr(71); Tab(T5); "NET AMOUNT: "; Tab(T8 + 5); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
            VNetamt = netamount
            Line = Line + 2
        End If
        kk.Close
        kk.Open "Select * from CASHA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(T5); kk!txt1; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + myround(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(T5); kk!txt2; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + myround(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(T5); "CASH RECD. "; Tab(T8 + 4); rsets(Trim(Format(str(Abs(myround(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - myround(kk!baa, 2)
             End If
        End If
        Print #1, Tab(T8); repli("-", 22)
        Print #1, Tab(T5); Chr(27) + Chr(71); "BALANCE : "; Tab(T8 + 5); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
        Line = Line + 2
        'PRINT THE FOOTER IN INVOICE START
        Do While Line < 65
            Print #1, ""
            Line = Line + 1
        Loop
        
        
        Print #1, Tab(0); toword(myround(VNetamt, 2))
        Print #1, Tab(0); repli("-", 150)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        Dim LEFTM As Integer
        LEFTM = 5
        
        CNSetup
        tempdata.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!cname)) * 0.65))); "FOR " + Trim(tempdata!cname)
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
Dim FooterYes As Boolean
Dim totalquantity As Long
Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
Dim rs As ADODB.Recordset
Dim LEFTM As Integer
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
Set kkk = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Open "" + App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
LEFTM = 5
FooterYes = False
header:
    If kkk.State = 1 Then
          kkk.Close
    End If
    CNSetup
    kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
     Print #1, Chr(27) + Chr(77)
     Line = Line + 7
End If
Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CASH MEMO")))) / 2 - 10); Chr(14); "***CASH MEMO***"; Chr(20); Tab(50); IIf(printheader = True, kkk!uptt, "")
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
rs1.Open "select * from casha where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
'Tab(20); Mid$(rs1!SUBLEDGER, 1, 5);
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); " To,"; Tab(7); IIf(Optioncash.Value = True, "", Mid$(rs1!subledger, 1, 5)); Tab(48); "Cash Memo No. : "; Trim(rs1!INVOICENO); Tab(82); "Dt. : "; rs1!InvoiceDate; Chr(27) + Chr(72);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from sledger where   " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & IIf(Optioncash.Value = True, rs1!CASHPARTYNAME, kkk!DESCFORINVOICE)
        Print #1, Tab(5); IIf(IsNull(kkk!address1), "", kkk!address1); Tab(49); Chr(27) + Chr(71); "Order by    : "; Chr(27) + Chr(72); Trim(rs1!ORDERBY); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!OrderDate), "  /  /    ", rs1!OrderDate)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS2), "", kkk!ADDRESS2); Tab(49); Chr(27) + Chr(71); "Bilty No.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!BILTYDATE), "  /  /    ", rs1!BILTYDATE)
        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS3), "", kkk!ADDRESS3)
        kkk.Close
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
    If kk.State = 1 Then kk.Close
    kk.Open "select * from CASHB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,sno ", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where " & stringyear & " and bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
                    
                    
printagain2:
                    called2 = False
                End If
                Print #1, Tab(70); repli("-", 12)
                Line = Line + 1
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CashB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
    If kk.State = 1 Then kk.Close
    kk.Open "Select * from CASHC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kk.BOF Then
       Do While Not kk.EOF
             If kk!amount > 0 Then
                    If Trim(kk!DEBITORCREDIT) = Trim("Credit") Then
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
        kk.Close
        kk.Open "Select * from CASHA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                 Print #1, Tab(59); Chr(27) + Chr(71); "BALANCE   : "; Tab(85); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
                 Print #1, Tab(84); repli("-", 12);
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
        Print #1, ""
        'PRINT THE FOOTER IN INVOICE END
        Close #1
        PrintOption.Show
        
End Sub
