VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBookIssueSp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7110
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmIssueSp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleMode       =   0  'User
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboGodown 
      Height          =   315
      Left            =   7290
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   945
      Width           =   660
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2925
      Left            =   10005
      TabIndex        =   55
      Top             =   1635
      Visible         =   0   'False
      Width           =   6315
      Begin VB.ComboBox customercode 
         Height          =   1155
         Left            =   315
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   65
         Top             =   735
         Visible         =   0   'False
         Width           =   3765
      End
      Begin MSMask.MaskEdBox textbox 
         Height          =   315
         Left            =   960
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_DTOB 
         Height          =   315
         Left            =   1650
         TabIndex        =   58
         Top             =   315
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_OB 
         Height          =   315
         Left            =   0
         TabIndex        =   59
         Top             =   315
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox through 
         Height          =   285
         Left            =   0
         TabIndex        =   62
         Top             =   300
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox through1 
         Height          =   285
         Left            =   3270
         TabIndex        =   63
         Top             =   300
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Through : "
         Height          =   285
         Left            =   60
         TabIndex        =   64
         Top             =   -120
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dated : "
         Height          =   285
         Left            =   1650
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Order By : "
         Height          =   285
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cust Code : "
         Height          =   315
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.ComboBox cmbtransportname 
      Height          =   315
      Left            =   1260
      TabIndex        =   7
      Top             =   945
      Width           =   1965
   End
   Begin VB.TextBox txtadst 
      Height          =   315
      Left            =   4950
      TabIndex        =   53
      Top             =   4950
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ComboBox cmbAgentName 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   3750
   End
   Begin VB.CommandButton Commandall 
      Caption         =   "All Books"
      Height          =   375
      Left            =   1050
      TabIndex        =   47
      Top             =   5190
      Width           =   945
   End
   Begin VB.CommandButton Commandother 
      Caption         =   "&End Part"
      Height          =   375
      Left            =   180
      TabIndex        =   23
      Top             =   5190
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2880
      Left            =   0
      TabIndex        =   40
      Top             =   1320
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   5080
      _Version        =   393216
      FillStyle       =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.ComboBox Bookname 
      Height          =   960
      Left            =   3450
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   35
      Top             =   2550
      Width           =   2295
   End
   Begin VB.ComboBox Bookcode 
      Height          =   1740
      ItemData        =   "frmIssueSp.frx":000C
      Left            =   2640
      List            =   "frmIssueSp.frx":000E
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   34
      Top             =   2595
      Width           =   2355
   End
   Begin VB.PictureBox Picture5 
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9495
      TabIndex        =   13
      Top             =   5655
      Width           =   9555
      Begin VB.CommandButton Commandprintnh 
         Caption         =   "N&HPrint"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6300
         TabIndex        =   50
         Top             =   0
         Width           =   885
      End
      Begin VB.CommandButton Commandadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1020
         TabIndex        =   0
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   375
         Left            =   8340
         TabIndex        =   20
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7320
         TabIndex        =   19
         Top             =   30
         Width           =   945
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5370
         TabIndex        =   18
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4485
         TabIndex        =   17
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   375
         Left            =   3630
         TabIndex        =   16
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Height          =   375
         Left            =   2760
         TabIndex        =   25
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1860
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Commandhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   800
      End
   End
   Begin MSMask.MaskEdBox bundles 
      Height          =   285
      Left            =   7530
      TabIndex        =   5
      Top             =   360
      Width           =   1950
      _ExtentX        =   3440
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
      Height          =   315
      Left            =   1110
      TabIndex        =   2
      Top             =   330
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tempmeb 
      Height          =   285
      Left            =   120
      TabIndex        =   36
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
      Left            =   1020
      TabIndex        =   37
      Top             =   4080
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
      Left            =   60
      TabIndex        =   38
      Top             =   4470
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
      Left            =   60
      TabIndex        =   1
      Top             =   330
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.ComboBox Genledger 
      Height          =   315
      Left            =   5865
      Sorted          =   -1  'True
      TabIndex        =   39
      Top             =   4935
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSMask.MaskEdBox marka 
      Height          =   285
      Left            =   5910
      TabIndex        =   4
      Top             =   360
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox weight 
      Height          =   315
      Left            =   7950
      TabIndex        =   12
      Top             =   960
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox freight 
      Height          =   315
      Left            =   5880
      TabIndex        =   10
      Top             =   960
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bdated 
      Height          =   315
      Left            =   4680
      TabIndex        =   9
      Top             =   960
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
      Left            =   3210
      TabIndex        =   8
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox station 
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Godown : "
      Height          =   285
      Left            =   7290
      TabIndex        =   66
      Top             =   675
      Width           =   645
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transport"
      Height          =   285
      Left            =   1260
      TabIndex        =   54
      Top             =   660
      Width           =   1935
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agent :"
      Height          =   315
      Left            =   2160
      TabIndex        =   52
      Top             =   30
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Press F4 Key To Delete A Invoive Item"
      Height          =   405
      Left            =   3030
      TabIndex        =   51
      Top             =   6135
      Width           =   3705
   End
   Begin VB.Label labelbybanklbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By Bank : "
      Height          =   255
      Left            =   2340
      TabIndex        =   49
      Top             =   5295
      Width           =   1200
   End
   Begin VB.Label labelbybank 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3540
      TabIndex        =   48
      Top             =   5295
      Width           =   1200
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   4695
      TabIndex        =   24
      Top             =   660
      Width           =   1185
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marka : "
      Height          =   285
      Left            =   5925
      TabIndex        =   29
      Top             =   60
      Width           =   1605
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Quantity : "
      Height          =   255
      Left            =   3300
      TabIndex        =   46
      Top             =   4605
      Width           =   1470
   End
   Begin VB.Label tqu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3300
      TabIndex        =   45
      Top             =   4920
      Width           =   1485
   End
   Begin VB.Label mgd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   7980
      TabIndex        =   43
      Top             =   4905
      Width           =   1200
   End
   Begin VB.Label mna 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   7980
      TabIndex        =   42
      Top             =   5265
      Width           =   1200
   End
   Begin VB.Label mga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6720
      TabIndex        =   41
      Top             =   4935
      Width           =   1200
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   1110
      TabIndex        =   33
      Top             =   15
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Challan No. : "
      Height          =   285
      Left            =   30
      TabIndex        =   32
      Top             =   15
      Width           =   1065
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Net Amount : "
      Height          =   255
      Left            =   6780
      TabIndex        =   31
      Top             =   5265
      Width           =   1200
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Gross Amount : "
      Height          =   255
      Left            =   6720
      TabIndex        =   30
      Top             =   4665
      Width           =   1260
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bundle(s) : "
      Height          =   285
      Left            =   7530
      TabIndex        =   28
      Top             =   60
      Width           =   1950
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Railway/Station : "
      Height          =   285
      Left            =   0
      TabIndex        =   27
      Top             =   660
      Width           =   1230
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bilty No. : "
      Height          =   285
      Left            =   3210
      TabIndex        =   26
      Top             =   660
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Freight : "
      Height          =   285
      Left            =   5880
      TabIndex        =   22
      Top             =   660
      Width           =   1410
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Weight : "
      Height          =   285
      Left            =   7950
      TabIndex        =   21
      Top             =   660
      Width           =   1545
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Discount : "
      Height          =   255
      Left            =   7980
      TabIndex        =   44
      Top             =   4635
      Width           =   1290
   End
   Begin VB.Menu dd 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmBookIssueSp"
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
        Print #1, Tab(1); kkk!COURT; Tab(60); "FOR " + Trim(kkk!CNAME)
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
rs1.Open "invoicea", CON, adOpenStatic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); "AGENT NAME : "; Chr(27) + Chr(71); Mid$(rs1!agentname, 1, 20); Tab(45); Chr(27) + Chr(71); "  Challan No. : "; Chr(27) + Chr(71); Trim(rs1!INVOICENO); Tab(75); Chr(27) + Chr(71); "  Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!invoicedate), "", rs1!invoicedate)
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
''''''    kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
''''''    If Not kkk.EOF Then
''''''        Print #1, Tab(5); "M/s " & kkk!DESCFORINVOICE
''''''        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS1), " ", kkk!ADDRESS1); Tab(49); Chr(27) + Chr(71); "Order by    : "; Chr(27) + Chr(72); Trim(rs1!ORDERBY); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
''''''        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS2), " ", kkk!ADDRESS2)
''''''        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS3), " ", kkk!ADDRESS3); Tab(49); Chr(27) + Chr(71); "Bilty No.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!biltydate), "  /  /    ", rs1!biltydate)
''''''        Print #1, ""
''''''        kkk.close
      Print #1, Tab(45); Chr(27) + Chr(71); "Bilty NO.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(75); Chr(27) + Chr(71); "Dt  : "; Chr(27) + Chr(72); IIf(IsNull(Trim(rs1!biltydate)), "", Trim(rs1!biltydate))
      Print #1, ""
      Print #1, ""
      Print #1, ""
      Print #1, Tab(0); Chr(27) + Chr(71); "(" & cbogodown & ")"; Chr(27) + Chr(72)
      Print #1, Chr(27) + Chr(71); "Station   :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!station); " "; Trim(rs1!transportname); Tab(75); Chr(27) + Chr(71); "Pvt. Mark : "; Chr(27) + Chr(72); Trim(rs1!marka)
      
'Print #1, Chr(27) + Chr(71); "Bilty NO :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!biltyno); Chr(27) + Chr(71); "  Bilty Date  : "; Chr(27) + Chr(72); Trim(rs1!biltydate)
      'Print #1, Chr(27) + Chr(71); "Freight  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!freight); Chr(27) + Chr(71); " Bilty NO :"; Chr(27) + Chr(72); Tab(40); Trim(rs1!biltyno); Tab(50); Chr(27) + Chr(71); "Weight  : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."; Tab(60); Chr(27) + Chr(71); "Bundle(s)  : "; Chr(27) + Chr(72); Trim(rs1!bundles)
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
    kk.Open "select * from invoiceb where invoiceno=" + Trim(rs1!INVOICENO) + " order by printorder,sno ", CON, adOpenStatic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
                Print #1, Tab(0); rsets(Trim(str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(str(kk!quantity)), 5); Tab(58); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from invoiceb where invoiceno=" + Trim(rs1!INVOICENO) + " and printorder =" + Trim(str(cdiscount)) + " group by printorder", CON, adOpenStatic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(str(tdata(0)), "0.00")), 12)
                   'Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(Str(tdata(0) * cdiscount / 100), "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata(0) - (tdata(0) * cdiscount / 100)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(str(vdis), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(70); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.close
             Loop
           End If
       End If
       Print #1, repli("-", 96)
       Print #1, Tab(50); rsets(Trim(str(totalquantity)), 7); Tab(84); rsets(Trim(Format(str(netamount), "0.00")), 12)
       Line = Line + 2
       If kk.State = 1 Then
             kk.close
       End If
       kk.Open "Select * from invoicec where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
       If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
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
        kk.close
        Dim Va As Variant
        kk.Open "Select * from invoicea where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
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
        Print #1, Tab(0); Chr(27) + Chr(71); toword(Round(VNetamt, 2)); Chr(27) + Chr(72)
        Print #1, Tab(0); repli("-", 96)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        CNSetup
        tempdata.Open "setup1", CON, adOpenStatic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(1); tempdata!COURT; Tab(60); "FOR " + Trim(tempdata!CNAME)
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
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
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
        Unload OTHERSALES
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
                    rs.close
                End If
                rs.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                row = Grid1.row
                col = Grid1.col
                If Trim(Grid1.Text) <> "" Then
                    If Not rs.BOF Then
                        rs.MoveFirst
                        rs.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If rs.EOF Then
                            tempmeb.Visible = True
                            tempmeb.SetFocus
                            rs.close
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
                                
                                
                             'If Edit = False Then
                                
                                Grid1.col = 5
                                If Grid1.Text = "" Then
                                Grid1.Text = Format(rs(3), "0.00")            'rs(3)
                                End If
                                r = rs(3)
                                
                                Grid1.col = 4
                                If Grid1.Text = "" Then
                                Grid1.Text = Format(rs(4), "0.00")
                                End If
                                
                                Grid1.col = 6
                                If Grid1.Text = "" Then
                                Grid1.Text = Format(rs(4), "0.00")
                                End If
                                
                                D = rs(4)
                            'End If
                            
                                Grid1.col = 7
                                Grid1.Text = Format(Round(q * r, 2), "0.00")
                                Grid1.col = 8
                                Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
                             ' Else
                              
                                  If Grid1.Text = "" And addmode = False Then
                                    If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        kk.close
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
                                        Grid1.col = 4
                                    '    If kk.BOF Then
                                     '        GoTo abc
                                     '   End If
                                        Grid1.Text = Format(kk(0), "0.00")
                                        Grid1.col = 6
                                        Grid1.Text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = rs(3)
                                    End If
                                  End If
  
  '                            End If
                          '  End If
                            Grid1.col = col
                            rs.close
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
            totalamount = totalamount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.Text)), 2)
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
                           Commandother.SetFocus
                           Exit Sub
                    End If
                End If
                Grid1.row = row
                Grid1.col = col
                Grid1.Text = Bookname.Text
                '/*************************
                If rs.State = 1 Then
                    rs.close
                End If
                rs.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                row = Grid1.row
                col = Grid1.col
                If Trim(Grid1.Text) <> "" Then
                    If Not rs.BOF Then
                        rs.MoveFirst
                        rs.Find "bookname='" + Trim(Grid1.Text) + "'"
                        If rs.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            rs.close
                            Exit Sub
                        Else
                            
                            Grid1.col = 1
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
                            Grid1.Text = Format(rs(3), "0.00")
                            r = rs(3)
                            '/******************
'''                            Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
'''                            Grid1.col = 6
'''                            If Trim(kk(0)) <> "" Then
'''                               tempstr = Trim(kk(0))
'''                               kk.Close
'''                               Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
'''                               Grid1.col = 4
'''                               If kk.BOF Then
'''                                   GoTo abc
'''                               End If
'''                               Grid1.Text = Format(kk(0), "0.00")
'''                               Grid1.col = 6
'''                               Grid1.Text = Format(kk(0), "0.00")
'''                               D = kk(0)
'''                            Else
'''abc:
                                 Grid1.col = 4
                                 Grid1.Text = Format(rs(4), "0.00")
                                    Grid1.col = 6
                                    Grid1.Text = Format(rs(4), "0.00")
                                    D = rs(4)
                      '          End If
                                Grid1.col = 7
                                Grid1.Text = Round(q * r, 2)
                                Grid1.col = 8
                                Grid1.Text = Round((q * r) * (D / 100), 2)
                         '   End If
                            Grid1.col = col
                            rs.close
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
            totalamount = totalamount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.Text)), 2)
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
   Exit Sub
Else
  Dim rs1 As New ADODB.Recordset
  rs1.Open "select *  from AgentMaster where AgentName='" & cmbAgentName.Text & "' order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
  If rs1.RecordCount <= 0 Then
     MsgBox "Enter valid Agent Name.. "
     cmbAgentName.SetFocus
  End If
End If

End Sub

Private Sub CommandPrint_LostFocus()
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub

Private Sub Commandprintnh_Click()
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
    Dim rs As ADODB.Recordset
    addoredit = True
    addmode = True
    Set rs = New ADODB.Recordset
    Dim TEMPNUM As Integer
    If Edit = False Then
       'If CON.Execute("Select max(invoiceno) from invoicea")(0) >= Val(Trim(Me.I_NO.Text)) Then
          Me.I_NO.Text = CON.Execute("Select max(invoiceno) from invoicea")(0) + 1
          rs.Open "tempinv", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
          If rs.BOF Then
             rs.AddNew
          End If
          Me.I_NO.Text = rs!In + 1
          rs!In = Val(Me.I_NO.Text)
          rs.update
          rs.close
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
    Me.customercode.Enabled = True
    cbogodown.ListIndex = 0
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
        rs.close
    End If
    rs.Open "select * from books order by BOOKCODE", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
            Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
            Grid1.col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.close
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
            Grid1.Text = Format(Round(q * r, 2), "0.00")
            Grid1.col = 8
            Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
            If Not rs.EOF Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.row = Grid1.row + 1
                rs.MoveNext
            End If
        Loop
        '/**fghfghgh
        '    Grid1.col = col
    End If
    rs.close
   ' row = Grid1.row
   ' col = Grid1.col
    totalamount = 0
    totaldiscount = 0
    Me.tqu.Caption = ""
    For I = 1 To Grid1.Rows - 1
            Grid1.row = I
            Grid1.col = 7
            totalamount = totalamount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.Text)), 2)
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

    
    
    
    
 '=====================================================================================

    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
   
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from invoicea where invoiceno=" & I_NO.Text & "", CON
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select * from invoicea where invoiceno=" & I_NO.Text & "", CON
       'If rs_h.Fields("Print_yes").Value = "y" Then
          If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
       '   End If
       End If
       
    End If
    
'======================================================================================
    
    
'===================================================================================
    
    
    If rs_v.State = 1 Then rs_v.close
    rs_v.Open "select BDelete,Bedit,Bsave from setup where UserId=" & UId & "", CONINFO
    If rs_v.EOF = False Then
    If rs_v!bDelete = False Then
       MsgBox "You Can'nt Delete ...", vbCritical
    Exit Sub
    End If
    End If


  
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
      

If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
                CON.Execute ("delete * from invoicea where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete * from invoiceb where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete * from invoicec where INVOICENO = " + Trim(I_NO.Text))
                invoiceabandon
End If
End Sub

Private Sub Commandedit_Click()
    
    If rs_v.State = 1 Then rs_v.close
    rs_v.Open "select BDelete,Bedit,Bsave from setup where UserId=" & UId & "", CONINFO
    If rs_v.EOF = False Then
    If rs_v!bedit = False Then
       MsgBox "You Can'nt Edit ...", vbCritical
    Exit Sub
    End If
    End If


    
    
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
    ' invoicectmp creation start
    CON.Execute ("delete * from invoicectmp WHERE INVOICENO = " + Trim(I_NO.Text))
    DoEvents
    CON.Execute ("insert into invoicectmp  select * from invoicec where INVOICENO = " + Trim(I_NO.Text))
    DoEvents
    ' invoicetmp creation end
    Dim kx As Integer
    kx = 0
    Do While kx < 18000
    kx = kx + 1
    Loop
    addoredit = False
    
    HIT
    VB.Screen.ActiveForm.ActiveControl.SelLength = 2
    
    OTHERSALES.Top = 0
    OTHERSALES.Left = 0
    OTHERSALES.Visible = False
End Sub
Private Sub Commandother_Click()
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    Commandsave.Enabled = True
    Me.Enabled = False
    OTHERSALES.Show
    OTHERSALES.Top = 0
    OTHERSALES.Left = 0
    'Unload OTHERSALES
    'Load OTHERSALES
    'OTHERSALES.Show
    'OTHERSALES.Top = 0
   ' OTHERSALES.Left = 0
  'OTHERSALES.Show
  
 
End Sub
Private Sub CommandPrint_Click()
  Printheader = True
  printinvoice
   
End Sub

Private Sub Commandprintnh_LostFocus()
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub

Private Sub Commandreturn_Click()
   Dim rs As New ADODB.Recordset
   rs.Open "tempINV", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
   If rs.BOF Then
       rs.AddNew
   End If
   rs!In = CON.Execute("Select max(invoiceno) from INVOICEA")(0)
   rs.update
   rs.close
   Unload Me
   addoredit = False
    
End Sub
Private Sub Commandsave_Click()
    
    
    
    
    
    
'=====================================================================================

    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
   
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from invoicea where invoiceno=" & I_NO.Text & "", CON
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select * from invoicea where invoiceno=" & I_NO.Text & "", CON
       'If rs_h.Fields("Print_yes").Value = "y" Then
          If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
       '   End If
       End If
       
    End If
    
'======================================================================================
    
    
    
    Dim SAVED As Boolean
    Dim LAMOUNT As Double
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If Edit = False And addmode = False Then
      Me.Commandsave.Enabled = False
      Exit Sub
    End If
    
    If rs_v.State = 1 Then rs_v.close
    rs_v.Open "select BDelete,Bedit,Bsave from setup where UserId=" & UId & "", CONINFO
    If rs_v.EOF = False Then
    If rs_v!bsave = False Then
       MsgBox "You Can'nt Save ...", vbCritical
    Exit Sub
    End If
    End If

   
  
    
    
    
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
    
    'If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(customercode.Text) <> "" Then
     If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(Me.cmbAgentName.Text) <> "" Then
            If Edit Then
                CON.Execute ("delete * from invoicea where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete * from invoiceb where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete * from invoicec where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete * from cashregister where CMNo = " + Trim(I_NO.Text))
            End If
            If rs.State = 1 Then
                rs.close
            End If
            LAMOUNT = 0
            rs.Open "select * from invoicea where invoiceno <=0", CON, adOpenDynamic, adLockOptimistic
            If Not Edit Then
again:
           If CON.Execute("Select max(invoiceno) from invoicea")(0) >= Val(Trim(Me.I_NO.Text)) Then
                   ' Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'GoTo again
                End If
            End If
            rs.AddNew
            rs!INVOICENO = Val(Me.I_NO.Text)
            rs!invoicedate = Me.i_dt.Text
            rs!Godown = cbogodown.Text
            'rs!Genledger = Trim(Me.Genledger.Text)
            'rs!SUBLEDGER = Trim(Me.customercode.Text)
            rs!agentname = Trim(Me.cmbAgentName.Text)
            rs!transportname = Trim(Me.cmbtransportname.Text)
            rs!orderby = Trim(Me.I_OB.Text)
            If Trim(Me.I_DTOB) <> Trim("__/__/____") Then
            '    rs!ORDERDATE = Date
            'Else
                rs!ORDERDATE = Trim(Me.I_DTOB.Text)
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
                rs!biltydate = Me.bdated & ""
           End If
            rs!freight = Trim(Me.freight)
            rs!weight = Trim(Me.weight)
            rs!netamount = Round(Val(Trim(Me.mna.Caption)), 2)
            rs!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
            rs!txt1 = Trim(OTHERSALES.T1TEXT.Text)
            rs!txt1a = Val(Trim(OTHERSALES.T1.Text))
            rs!txt2 = Trim(OTHERSALES.T2TEXT.Text)
            rs!txt2a = Val(Trim(OTHERSALES.T2.Text))
            rs!baa = Val(Trim(OTHERSALES.T3TEXT.Text))
            rs!baa = Val(Trim(labelbybank.Caption))
            If addmode = True Then
                If Val(Trim(OTHERSALES.T3TEXT.Text)) <> 0 Then
                      rs!advicestatus = "Pending"
                      Me.txtadst.Text = "Pending"
                End If
            Else
                rs!advicestatus = Me.txtadst.Text & ""
            End If
            Dim trs As New ADODB.Recordset
            trs.Open " SELECT DISTCODE FROM SLEDGER  WHERE SUBLEDGER='" & customercode.Text & "'", CON, adOpenStatic, adLockOptimistic, adCmdText
            If Not trs.BOF Then
                rs!District = Trim(trs!distcode)
            Else
                rs!District = ""
            End If
err1:
           If Not Edit Then
                If CON.Execute("Select max(invoiceno) from invoicea")(0) >= Val(Trim(Me.I_NO.Text)) Then
                    'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'rs!INVOICENO = Val(Me.I_NO.Text)
                    On Error GoTo err1
                End If
            End If
            
            rs.update
            On Error GoTo 0
            rs.close
            rs.Open "select * from invoiceb where invoiceno<=0", CON, adOpenDynamic, adLockOptimistic
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
                         rs.AddNew
                         Grid1.col = 1
                         rs!INVOICENO = Val(Me.I_NO.Text)
                         rs!invoicedate = Me.i_dt.Text
             '            rs!Genledger = Trim(Me.Genledger.Text)
              '           rs!SUBLEDGER = Trim(Me.customercode.Text)
                         rs!Bookcode = Trim(Grid1.Text)
                         Grid1.col = 3
                         rs!quantity = Trim(Grid1.Text)
                         Grid1.col = 5
                         rs!rate = Trim(Grid1.Text)
                         Grid1.col = 7
                         rs!amount = Trim(Grid1.Text)
                         LAMOUNT = Val(Trim(Grid1.Text))
                         Grid1.col = 4
                         rs!PRINTORDER = Trim(Grid1.Text)
                         Grid1.col = 6
                         rs!discount = Trim(Grid1.Text)
                         Grid1.col = 8
                         rs!netamount = LAMOUNT - Trim(Grid1.Text)
                         LAMOUNT = 0
                         rs!agentname = Trim(Me.cmbAgentName.Text)
                         rs.update
                       End If
                    End If
                End If
            Next
            rs.close
            Grid1.TopRow = 1
            Grid1.row = 1
            Grid1.col = 1
            rs.Open "select * from invoicec where invoiceno<=0", CON, adOpenDynamic, adLockOptimistic
            '/******
                'Dim I, x As Integer
                Dim temprs As ADODB.Recordset
                Set temprs = New ADODB.Recordset
                For I = 1 To OTHERSALES.mrow
                    OTHERSALES.Grid1.row = I
                    OTHERSALES.Grid1.col = 0
                    If Trim(OTHERSALES.Grid1.Text) <> "" Then
                        rs.AddNew
                        rs!INVOICENO = Val(Me.I_NO.Text)
                        rs!invoicedate = Me.i_dt.Text
                        rs!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
                        rs!Text = Trim(OTHERSALES.Grid1.Text)
                        If temprs.State = 1 Then
                            temprs.close
                        End If
                        If Edit Then
                        temprs.Open "select * from INVOICEctmp WHERE INVOICENO=" & INVOICE.I_NO & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
                        If OTHERSALES.Grid1.Text <> "" Then
                                temprs.Find "TEXT='" + Trim(OTHERSALES.Grid1.Text) + "'"
               '                 rs!Genledger = Trim(temprs!Genledger)
                '                rs!SUBLEDGER = Trim(temprs!SUBLEDGER)
                                rs!DebitorCredit = Trim(temprs!DebitorCredit)
                                rs!RYN = temprs!RYN & ""
                                
                        End If
                        temprs.close
                        
                        Else
                        
                        temprs.Open "select * from INVOICEEND  order by printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
                        If OTHERSALES.Grid1.Text <> "" Then
                                temprs.Find "TEXT='" + Trim(OTHERSALES.Grid1.Text) + "'"
                 '               rs!Genledger = Trim(temprs!Genledger)
                  '              rs!SUBLEDGER = Trim(temprs!SUBLEDGER)
                                rs!DebitorCredit = Trim(temprs!DebitorCredit)
                                rs!RYN = temprs!RYN & ""
                        End If
                        temprs.close
                        End If
                        OTHERSALES.Grid1.col = 1
                        rs!rate = Val(Trim(OTHERSALES.Grid1.Text))
                        If Val(Trim(OTHERSALES.Grid1.Text)) > 0 Then
                            rs!amount = Round((Me.totalamount - Me.totaldiscount), 2) * Round((Val(Trim(OTHERSALES.Grid1.Text)) / 100), 2)
                        Else
                        OTHERSALES.Grid1.col = 2
                            rs!amount = Val(Trim(OTHERSALES.Grid1.Text))
                        End If
                    rs.update
                    End If
                Next
                rs.close
                CON.Execute ("delete * from INVOICECtmp where INVOICENO = " + Trim(I_NO.Text))
                 
                rs.Open "tempINV", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
                If rs.BOF Then
                    rs.AddNew
                End If
                rs!In = CON.Execute("Select max(invoiceno) from INVOICEA")(0)
                rs.update
                rs.close
                
                
              If Me.station.Text <> "" Then
                    
                    S11 = ""
                    ss11 = ""
                    
                    S11 = InStr(1, Me.station.Text, " ")
                    If S11 <> 0 Then
                    ss11 = Trim(Mid(Me.station.Text, 1, S11))
                    Else
                    ss11 = Me.station.Text
                    End If
                    PopUpValue1 = ss11

                 UpdateDisPatchReg1 I_NO, i_dt, Me.cmbAgentName.Text, PopUpValue1, Trim(Me.bundles), Trim(Me.cmbtransportname.Text), "-", Trim(Me.biltno.Text), Me.bdated, Trim(Me.freight), "CashRegister"
                 PopUpValue1 = ""
              End If

                
                
                
''                If addmode = True Then
''                    rs.Open "tempinv", CON1, adOpenKeyset, adLockOptimistic, adCmdTable
''                    If rs.BOF Then
''                        rs.AddNew
''                    End If
''                    rs!In = Val(Me.I_NO.Text)
''                    rs.Update
''                    rs.Close
''                End If
            SAVED = True
        End If
        If SAVED Then
            MsgBox "Record Saved"
            Unload OTHERSALES
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
    On Error Resume Next
    Unload OTHERSALES
    CON.Execute "select * from INVOICEctmp WHERE INVOICENO=" & INVOICE.I_NO & " "
    Me.Enabled = False
    'searchscreen.Grid1.row = 0
    'searchscreen.Grid1.col = 0
    Call searchscreen.tempr(11, "invoice")
End Sub

Private Sub customercode_LostFocus()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select * from sledger where gledger='SUNDRY DEBTORS' and subledger='" + Trim(customercode.Text) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If rs.RecordCount <= 0 Then
        customercode.SetFocus
        HIT
        rs.close
        Exit Sub
    End If
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    If rs!distcode <> "" And addmode = True Then
       rs1.Open "Select * from Districts where Districtname = '" & rs!distcode & "'", CON, adOpenStatic, adLockReadOnly
       If rs1.RecordCount > 0 Then
          Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
       End If
    End If
    rs.close
    Me.textbox.Text = Me.customercode.Text
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

Private Sub Form_Activate()
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
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
    Grid1.Left = 90
    Grid1.Top = 1500
    'Set CON = New ADODB.Connection
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
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
    Grid1.ColWidth(0) = 200
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 2500
    Grid1.ColWidth(3) = 750
    Grid1.ColWidth(4) = 750
    Grid1.ColWidth(5) = 850
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Bookname.Height = 2325
    Me.CommandPrint.Enabled = True
    Me.Commandprintnh.Enabled = True
    rs.Open "select * from books", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.Bookcode.AddItem rs(0)
            Me.Bookname.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.close
    Genledger.Text = "SUNDRY DEBTORS"
    rs.Open "select * from sledger where gledger='" + Trim(Genledger.Text) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.customercode.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.close
     '*******Agent  combo fill
    rs.Open "select  Agentname from AgentMaster order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
    cmbAgentName.Clear
    If Not rs.EOF Then
       Do While Not rs.EOF
          If IsNull(rs(0)) = False Then
            Me.cmbAgentName.AddItem rs(0)
          End If
          If Not rs.EOF Then rs.MoveNext
        Loop
    End If
    rs.close
    
    
    rs.Open "select transportname from transportMaster order by transportname", CON, adOpenDynamic, adLockReadOnly, adCmdText
    cmbtransportname.Clear
    If Not rs.EOF Then
       Do While Not rs.EOF
          If IsNull(rs(0)) = False Then
            Me.cmbtransportname.AddItem rs(0)
          End If
          If Not rs.EOF Then rs.MoveNext
        Loop
    End If
    rs.close

    
    
    
    rs.Open "select Godwn from godownMaster order by id"
    While rs.EOF = False
          cbogodown.AddItem rs(0)
          rs.MoveNext
    Wend
    rs.close
    
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
  
    rs.Open "tempinv", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
    If Not rs.BOF Then
        Me.I_NO.Text = rs!In
        INVOICE.Enabled = True
        INVOICE.Edit = False
        INVOICE.I_NO_LostFocus
        INVOICE.I_NO.Enabled = False
        lastrow = 0
        lastcol = 1
        Dim ctl As Control
        For Each ctl In INVOICE.Controls
            If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = False
            End If
            If UCase(Trim(ctl.Name)) = UCase(Trim(INVOICE.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(INVOICE.Commandall.Name)) Then
                ctl.Enabled = False
            End If
        Next
        INVOICE.Picture5.Enabled = True
        addoredit = False
        SendKeys "{TAB}"
    Else
       kk.Open "SELECT MAX(INVOICENO) FROM INVOICEA", CON, adOpenDynamic, adLockReadOnly, adCmdText
       If kk(0) <> "" Then
          Me.I_NO.Text = Trim(str(kk(0) + 1))
       Else
          Me.I_NO.Text = "1"
       End If
       kk.close
    End If
    rs.close
    
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
    
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub
Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub

Private Sub Grid1_Click()
'If Trim(Me.customercode.Text) <> "" Then
If Trim(Me.cmbAgentName.Text) <> "" Then


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
        
''        If Trim(Me.customercode.Text) <> "" Then
''            If Me.customercode.Enabled = True Then
''                Me.customercode.Enabled = False
''            End If
            
If Trim(Me.cmbAgentName.Text) <> "" Then
            If Me.cmbAgentName.Enabled = True Then
                Me.cmbAgentName.Enabled = False
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
    
'''    If Trim(Me.customercode.Text) <> "" Then
'''        If Me.customercode.Enabled = True Then
'''            Me.customercode.Enabled = False
'''        End If
        
        If Trim(Me.cmbAgentName.Text) <> "" Then
        If Me.cmbAgentName.Enabled = True Then
            Me.cmbAgentName.Enabled = False
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

Private Sub Grid1_KeyPress(KeyAscii As Integer)
'If Trim(Me.customercode.Text) <> "" Then
If Trim(Me.cmbAgentName.Text) <> "" Then

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
    trs2.Open "Select invoiceno as cn from invoicea", CON, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount <= 0 Then
       Exit Sub
    Else
'''''''        If tRS1.State = 1 Then tRS1.close
'''''''        tRS1.Open "Select min(invoiceno) as mid,invoicedate from  invoicea  group by invoiceno,invoiceDate", CON, adOpenDynamic, adLockOptimistic
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
    trs2.Open "Select max(invoiceno) as mid from invoicea where  invoicedate <= cdate('" & i_dt.Text & "')-1", CON, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount > 0 Then
    
        If IsNull(trs2!Mid) <> True Then
            If Val(I_NO.Text) >= trs2!Mid Then
               If tRS1.State = 1 Then tRS1.close
               tRS1.Open "Select  min(InvoiceNo)as m2 from invoicea where invoicedate >= cdate('" & i_dt.Text & "')+1", CON, adOpenDynamic, adLockOptimistic
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
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    
    
    If Val(inviceNo) > 0 Then
    I_NO.Text = inviceNo
      ' cmdButto
    End If
    
    inviceNo = ""

    
    
    
    If Trim(I_NO.Text) = "" Then
        MsgBox "Invoice cannot be null"
        I_NO.SetFocus
    Else
        If rs.State = 1 Then
           rs.close
        End If
        ''rs.Open "INVOICEA", con, adOpenDynamic, adLockReadOnly, adCmdTable
        rs.Open "Select * from  INVOICEA where INVOICENO = " + Trim(I_NO.Text) + "", CON, adOpenStatic, adLockReadOnly
        ''rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
        If rs.EOF Then
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
        
        Dim ctl As Control
        For Each ctl In Me.Controls
            If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = True
            End If
        Next
        
        I_NO.Text = rs!INVOICENO
        Me.i_dt.Text = rs!invoicedate
        cbogodown.Text = rs!Godown & ""
        Me.Genledger.Text = Trim(rs!Genledger)
        Me.customercode.Text = Trim(rs!SUBLEDGER)
        Me.cmbAgentName.Text = IIf(IsNull(rs!agentname), "", rs!agentname)
        Me.cmbtransportname.Text = IIf(IsNull(rs!transportname), "", rs!transportname)
        Me.textbox.Text = Trim(rs!SUBLEDGER)
        Me.I_OB.Text = IIf(IsNull(rs!orderby), "", Trim(rs!orderby))
        If rs!ORDERDATE <> "" Then
        Me.I_DTOB.Text = rs!ORDERDATE
        End If
        Me.marka.Text = IIf(IsNull(rs!marka), "", Trim(rs!marka))
        Me.bundles = IIf(IsNull(rs!bundles), "", rs!bundles)
        Me.through.Text = IIf(IsNull(rs!through), "", rs!through)
        Me.through1.Text = IIf(IsNull(rs!through1), "", rs!through1)
        Me.station.Text = IIf(IsNull(rs!station), "", rs!station)
        Me.biltno.Text = IIf(IsNull(rs!biltyno), "", rs!biltyno)
        If rs!biltydate <> "" Then
        Me.bdated = rs!biltydate
        End If
        Me.freight = IIf(IsNull(rs!freight), "", rs!freight)
        Me.weight = IIf(IsNull(rs!weight), "", rs!weight)
       'Me.labelbybank = round(val(Trim(rs!baa)
        Me.labelbybank = Format(Round(rs!baa, 2), "0.00")
       ' mna.Caption = rs!netamount
        mna.Caption = Format(Round(rs!netamount, 2), "0.00")
       'Me.Combosldistrictcode.Text = IIf(IsNull(rs!district), "", rs!district)
        Me.txtadst = IIf(IsNull(rs!advicestatus), "", rs!advicestatus)
        rs.close
       
       ' OTHERSALES.Form_Load
'*/**/*/*/*/*//*/*
        If rs.State = 1 Then
                rs.close
        End If
'Commandedit.Enabled = True

       ' Unload OTHERSALES
        CON.Execute "select * from INVOICEctmp WHERE INVOICENO=" & INVOICE.I_NO & ""
       ' OTHERSALES.Form_Load
       ' 'rs.Open "INVOICEB", con, adOpenDynamic, adLockReadOnly, adCmdTable
       rs.Open "Select * from INVOICEB where INVOICENO =" + Trim(I_NO.Text) + " order by SNO", CON, adOpenStatic, adLockReadOnly
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
                    kk.close
                End If
                kk.Open "select * from books where bookcode='" + Trim(rs!Bookcode) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
                Grid1.col = 2
                Grid1.Text = Trim(kk!Bookname)
                Grid1.col = 3
                Grid1.Text = Trim(rs!quantity)
                Grid1.col = 5
                Grid1.Text = Format(Round(rs!rate, 2), "0.00")
                Grid1.col = 7
                Grid1.Text = Format(Round(rs!amount, 2), "0.00")
                Grid1.col = 4
                
                Grid1.Text = Format(Round(rs!PRINTORDER, 2), "0.00")
                Grid1.col = 6
                
                Grid1.Text = Format(Round(rs!discount, 2), "0.00")
                Grid1.col = 8
                Grid1.Text = Format(Round(rs!amount * (rs!discount / 100), 2), "0.00")
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
        For I = 1 To maxrow
            Grid1.row = I
            Grid1.col = 7
            totalamount = totalamount + Round(Val(Trim(Grid1.Text)), 2)
            Grid1.col = 8
            totaldiscount = totaldiscount + Round(Val(Trim(Grid1.Text)), 2)
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
                        rs.close
                    End If
                    rs.Open "books", CON, adOpenStatic, adLockReadOnly, adCmdTable
                    If Not rs.BOF Then
                        rs.MoveFirst
                        rs.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If rs.EOF And Trim(Grid1.Text) <> "" Then
                            rs.close
                            Exit Sub
                        Else
                            rs.close
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
                      MsgBox "Discount And Printorder Not Match.."
                      
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
        If Trim(Me.cmbAgentName.Text) <> "" Then
            Me.Grid1.SetFocus
            Grid1.col = 1
            Grid1.row = 1
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
             Print #1, Tab(((paperWidth - (Len(Trim(kkk!CNAME)) * 2)) / 2)); Chr(27) + Chr(18) + Chr(14); Trim(kkk!CNAME)
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
    rs1.Open "invoicea", CON, adOpenDynamic, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
   
    Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!SUBLEDGER; Tab(T5); "Invoice No. : "; Trim(rs1!INVOICENO); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!invoicedate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.close
            End If
            kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE;
                Print #1, Tab(3); kkk!ADDRESS1; Tab(T5); "Order by    : "; Trim(rs1!orderby); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!ORDERDATE
                Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.   : "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!biltydate
                
                
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
            kk.Open "select * from invoiceb where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
                        tdata.Open "select sum(amount) from invoiceb where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
           Print #1, Tab(T5 - 4); rsets(Trim(str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12)
           Print #1, Tab(T6); repli("-", 22)
           If kk.State = 1 Then
                kk.close
           End If
           kk.Open "Select * from invoicec where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
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
           kk.Open "Select * from invoicea where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
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
   Print #1, Tab(1); kkk!COURT; Tab(LEFTM + (paperWidth - ((Len(kkk!COURT) + Len(kkk!CNAME)) * 0.75))); "FOR " + Trim(kkk!CNAME)
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
     Print #1, Tab(((paperWidth - (Len(Trim(kkk!CNAME)) * 2)) / 2)); Chr(27) + Chr(15) + Chr(14); Trim(kkk!CNAME)
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
rs1.Open "invoicea", CON, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To, S.L. Code : "; Tab(T1 + 10); Mid$(rs1!SUBLEDGER, 1, 5); Tab(T5); "Invoice No. : "; Trim(rs1!INVOICENO); Tab(T8); "Dated     : "; rs1!invoicedate   'Chr(27) + Chr(18);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE
        Print #1, Tab(3); kkk!ADDRESS1; Tab(T5); "Order by    : "; Trim(rs1!orderby); Tab(T8); "Dated     : "; IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
        Print #1, Tab(3); kkk!ADDRESS2; Tab(T5); "Bilty No.   : "; Trim(rs1!biltyno); Tab(T8); "Dated     : "; IIf(IsNull(rs1!biltydate), "  /  /    ", rs1!biltydate)
        Print #1, Tab(3); kkk!ADDRESS3
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
    kk.Open "select * from invoiceb where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "select sum(amount) from invoiceb where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                    Print #1, Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0), 2)), "0.00")), 12)
                    Print #1, Tab(T5); "Less Discount @ " + Trim(Format(str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
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
        Print #1, Tab(T5 - 4); rsets(Trim(str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12)
       'by vk Print #1, Tab(t8); repli("-", 22)
        Line = Line + 2
        If kk.State = 1 Then
             kk.close
        End If
        kk.Open "Select * from invoicec where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5 + 21); Trim(kk!Text) + " :  @  " + Trim(Format(str(Round(kk!rate, 2)), "0.00")) & " % "; Tab(T8 + 5); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5 + 20); Trim(kk!Text) & " :"; Tab(T8 + 5); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T8); repli("-", 22)
            Print #1, Chr(27) + Chr(71); Tab(T5 + 20); "NET AMOUNT  : "; Tab(T8 + 6); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
            VNetamt = netamount
        Line = Line + 2
        End If
        kk.close
        kk.Open "Select * from invoicea where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(T5 + 20); kk!txt1 & "  :"; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + Round(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(T5 + 20); kk!txt2 & " :"; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - Round(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(T5 + 20); "BY BANK       :"; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - Round(kk!baa, 2)
             End If
        End If
        Print #1, Tab(T8); repli("-", 22)
        Print #1, Tab(T5 + 20); Chr(27) + Chr(71); "BALANCE    : "; Tab(T8 + 6); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
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
        Print #1, Tab(1); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!CNAME)) * 0.75))); "FOR " + Trim(tempdata!CNAME)
    
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

