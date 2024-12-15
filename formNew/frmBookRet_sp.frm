VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBookRet_sp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6390
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBookRet_sp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboGodown 
      Height          =   315
      Left            =   4140
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1035
      Width           =   750
   End
   Begin VB.ComboBox cmbAgentName 
      Height          =   315
      Left            =   6075
      TabIndex        =   5
      Top             =   30
      Width           =   3300
   End
   Begin VB.CommandButton Commandall 
      Caption         =   "All Books"
      Height          =   375
      Left            =   1050
      TabIndex        =   47
      Top             =   4725
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Commandother 
      Caption         =   "&End Part"
      Height          =   375
      Left            =   210
      TabIndex        =   23
      Top             =   4725
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2805
      Left            =   45
      TabIndex        =   40
      Top             =   1380
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   4948
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
      Height          =   1935
      Left            =   480
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   34
      Top             =   2040
      Width           =   2355
   End
   Begin VB.PictureBox Picture5 
      Height          =   465
      Left            =   165
      ScaleHeight     =   405
      ScaleWidth      =   8655
      TabIndex        =   13
      Top             =   5505
      Width           =   8715
      Begin VB.CommandButton Command1 
         Caption         =   "N&HPrint"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6015
         TabIndex        =   50
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   885
         TabIndex        =   0
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   375
         Left            =   7710
         TabIndex        =   20
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6870
         TabIndex        =   19
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5145
         TabIndex        =   18
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4260
         TabIndex        =   17
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   375
         Left            =   3435
         TabIndex        =   16
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Height          =   375
         Left            =   2565
         TabIndex        =   25
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1725
         TabIndex        =   15
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   45
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   800
      End
   End
   Begin MSMask.MaskEdBox weight 
      Height          =   285
      Left            =   4905
      TabIndex        =   12
      Top             =   1050
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox freight 
      Height          =   285
      Left            =   60
      TabIndex        =   8
      Top             =   1050
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bdated 
      Height          =   315
      Left            =   3765
      TabIndex        =   4
      Top             =   315
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox biltno 
      Height          =   315
      Left            =   2565
      TabIndex        =   3
      Top             =   315
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bundles 
      Height          =   285
      Left            =   2700
      TabIndex        =   10
      Top             =   1050
      Width           =   1410
      _ExtentX        =   2487
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
      Left            =   1485
      TabIndex        =   2
      Top             =   315
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tempmeb 
      Height          =   285
      Left            =   510
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
      Left            =   510
      TabIndex        =   37
      Top             =   3570
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
      Left            =   480
      TabIndex        =   38
      Top             =   3900
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
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.ComboBox customercode 
      Height          =   1740
      Left            =   6345
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.ComboBox Genledger 
      Height          =   315
      Left            =   5490
      Sorted          =   -1  'True
      TabIndex        =   39
      Top             =   4305
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox marka 
      Height          =   285
      Left            =   1650
      TabIndex        =   9
      Top             =   1050
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox textbox 
      Height          =   315
      Left            =   8535
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Godown : "
      Height          =   285
      Left            =   4140
      TabIndex        =   52
      Top             =   765
      Width           =   690
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agent :"
      Height          =   315
      Left            =   5025
      TabIndex        =   51
      Top             =   15
      Width           =   1020
   End
   Begin VB.Label labelbybanklbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By Bank : "
      Height          =   255
      Left            =   2340
      TabIndex        =   49
      Top             =   5010
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label labelbybank 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   48
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   3765
      TabIndex        =   24
      Top             =   15
      Width           =   1200
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dem."
      Height          =   285
      Left            =   1650
      TabIndex        =   28
      Top             =   750
      Width           =   1065
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Quantity : "
      Height          =   255
      Left            =   3450
      TabIndex        =   46
      Top             =   4320
      Width           =   1470
   End
   Begin VB.Label tqu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3420
      TabIndex        =   45
      Top             =   4590
      Width           =   1500
   End
   Begin VB.Label mgd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8010
      TabIndex        =   43
      Top             =   4635
      Width           =   1170
   End
   Begin VB.Label mna 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8010
      TabIndex        =   42
      Top             =   4905
      Width           =   1170
   End
   Begin VB.Label mga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6720
      TabIndex        =   41
      Top             =   4635
      Width           =   1260
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   1485
      TabIndex        =   33
      Top             =   0
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Book Return No. : "
      Height          =   285
      Left            =   60
      TabIndex        =   32
      Top             =   15
      Width           =   1395
   End
   Begin VB.Label label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cust Code : "
      Height          =   285
      Left            =   7440
      TabIndex        =   31
      Top             =   390
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Net Amount : "
      Height          =   255
      Left            =   6720
      TabIndex        =   30
      Top             =   4905
      Width           =   1260
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Gross Amount : "
      Height          =   255
      Left            =   6720
      TabIndex        =   29
      Top             =   4335
      Width           =   1260
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bundle(s) : "
      Height          =   285
      Left            =   2700
      TabIndex        =   27
      Top             =   750
      Width           =   1425
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bilty No. : "
      Height          =   285
      Left            =   2565
      TabIndex        =   26
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Station : "
      Height          =   285
      Left            =   60
      TabIndex        =   22
      Top             =   750
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Weight : "
      Height          =   285
      Left            =   4905
      TabIndex        =   21
      Top             =   750
      Width           =   1410
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Discount : "
      Height          =   255
      Left            =   8010
      TabIndex        =   44
      Top             =   4335
      Width           =   1170
   End
   Begin VB.Menu dd 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmBookRet_sp"
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
Me.Commandedit.Enabled = True
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = True
Me.Commandabandon.Enabled = True
Me.CommandPrint.Enabled = True
Me.Command1.Enabled = True
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
          kkk.Close
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

Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("SPECIMEN RETURN CHALLAN")))) / 2 - 3); Chr(14); "SPECIMEN RETURN CHALLAN"; Chr(20)
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
If rs1.State = 1 Then rs1.Close
rs1.Open "credita", CON, adOpenStatic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); "AGENT NAME : "; Chr(27) + Chr(71); Mid$(rs1!agentname, 1, 20); Tab(45); Chr(27) + Chr(71); "  Return No. : "; Chr(27) + Chr(71); Trim(rs1!INVOICENO); Tab(75); Chr(27) + Chr(71); "  Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!invoicedate), "", rs1!invoicedate)
    Line = Line + 1
                    Print #1, ""
                        Line = Line + 1
    
    If kkk.State = 1 Then
        kkk.Close
    End If
''''''    kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
''''''    If Not kkk.EOF Then
''''''        Print #1, Tab(5); "M/s " & kkk!DESCFORINVOICE
''''''        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS1), " ", kkk!ADDRESS1); Tab(49); Chr(27) + Chr(71); "Order by    : "; Chr(27) + Chr(72); Trim(rs1!ORDERBY); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
''''''        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS2), " ", kkk!ADDRESS2)
''''''        Print #1, Tab(5); IIf(IsNull(kkk!ADDRESS3), " ", kkk!ADDRESS3); Tab(49); Chr(27) + Chr(71); "Bilty No.   : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(83); Chr(27) + Chr(71); "Dt. : "; Chr(27) + Chr(72); IIf(IsNull(rs1!biltydate), "  /  /    ", rs1!biltydate)
''''''        Print #1, ""
''''''        kkk.close
      Print #1, Tab(45); Chr(27) + Chr(71); "Bilty NO.  : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(75); Chr(27) + Chr(71); "Dt  : "; Chr(27) + Chr(72); IIf(IsNull(Trim(rs1!biltydate)), "", Trim(rs1!biltydate))
      Print #1, ""
      Print #1, ""
      Print #1, Tab(0); Chr(27) + Chr(71); "(" & cboGodown & ")"; Chr(27) + Chr(72)
      '''''''''''''''''''''''''''''''''''  Print #1, Chr(27) + Chr(71); "Station   :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!station); " "; Trim(rs1!transportname); Tab(75); Chr(27) + Chr(71); "Pvt. Mark : "; Chr(27) + Chr(72); Trim(rs1!marka)
      
'Print #1, Chr(27) + Chr(71); "Bilty NO :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!biltyno); Chr(27) + Chr(71); "  Bilty Date  : "; Chr(27) + Chr(72); Trim(rs1!biltydate)
      'Print #1, Chr(27) + Chr(71); "Freight  :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!freight); Chr(27) + Chr(71); " Bilty NO :"; Chr(27) + Chr(72); Tab(40); Trim(rs1!biltyno); Tab(50); Chr(27) + Chr(71); "Weight  : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."; Tab(60); Chr(27) + Chr(71); "Bundle(s)  : "; Chr(27) + Chr(72); Trim(rs1!bundles)
       Print #1, Chr(27) + Chr(71); "Station   :"; Chr(27) + Chr(72); Tab(15); Trim(rs1!freight); Tab(40); Chr(27) + Chr(71); "Weight  : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."; Tab(67); Chr(27) + Chr(71); "Bundle(s)  : "; Chr(27) + Chr(72); Trim(rs1!bundles)
       Print #1, Chr(27) + Chr(71); repli("-", 96)
        Print #1, Tab(0); "S.No."; Tab(15); "Book Description"; Tab(50); "Quantity"; Tab(62); "Rate"; Tab(74); "Amount"; Tab(86); "Net Amount"
        Print #1, repli("-", 96); Chr(27) + Chr(72)
        'Line = Line + 10
            Line = Line + 9
''''''    End If
    If called1 Then
        GoTo printagain1
    End If
    If called2 Then
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.Close
    kk.Open "select * from creditb where invoiceno=" + Trim(rs1!INVOICENO) + " order by printorder,sno ", CON, adOpenStatic, adLockReadOnly, adCmdText
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from creditb where invoiceno=" + Trim(rs1!INVOICENO) + " and printorder =" + Trim(str(cdiscount)) + " group by printorder", CON, adOpenStatic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(str(tdata(0)), "0.00")), 12)
                   'Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(Str(tdata(0) * cdiscount / 100), "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata(0) - (tdata(0) * cdiscount / 100)), "0.00")), 12)
                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(str(vdis), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(70); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.Close
             Loop
           End If
       End If
       Print #1, repli("-", 96)
       Print #1, Tab(50); rsets(Trim(str(totalquantity)), 7); Tab(84); rsets(Trim(Format(str(netamount), "0.00")), 12)
       Line = Line + 2
       If kk.State = 1 Then
             kk.Close
       End If
       kk.Open "Select * from creditc where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
       If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount - kk!amount
                    Else
                        netamount = netamount + kk!amount
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
        kk.Open "Select * from credita where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
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


''''''
''''''Me.Commandadd.Enabled = True
''''''Me.Commandedit.Enabled = True
''''''Me.Commandsearch.Enabled = True
''''''Me.Commandsave.Enabled = False
''''''Me.Commanddelete.Enabled = True
''''''Me.Commandabandon.Enabled = True
''''''Me.CommandPrint.Enabled = True
''''''Me.Command1.Enabled = True
''''''Dim called1, called2 As Boolean
''''''Dim MaxLine As Integer
''''''Dim netamount As Double
''''''Dim totalquantity As Long
''''''Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
''''''Dim RS As ADODB.Recordset
''''''Dim Pno As Integer
''''''Set RS = New ADODB.Recordset
''''''T1 = 10
''''''T2 = 25
''''''T3 = 40
''''''T4 = 55
''''''T5 = 70
''''''T6 = 85
''''''T7 = 100
''''''T8 = 115
''''''netamount = 0
''''''totalquantity = 0
''''''paperWidth = 96
''''''MaxLine = 60
''''''called1 = False
''''''called2 = False
''''''Dim Line As Integer
''''''Dim rs1 As ADODB.Recordset
''''''Dim kkk As ADODB.Recordset
''''''Dim FooterYes As Boolean
''''''Set kkk = New ADODB.Recordset
''''''Set rs1 = New ADODB.Recordset
''''''Dim LEFTM As Integer
''''''Open "" + VB.App.Path + "\vipin.txt" For Output As #1
''''''Line = 0
''''''Pno = 1
''''''FooterYes = False
''''''header:
''''''If kkk.State = 1 Then kkk.close
''''''    CNSetup
''''''    kkk.Open "select * from setup1", CON, adOpenDynamic, adLockReadOnly, adCmdText
''''''    If FooterYes = True Then
''''''        If Line > MaxLine - 5 Then
''''''            Do While Line < 61
''''''                Print #1, ""
''''''                Line = Line + 1
''''''            Loop
''''''        End If
''''''        FooterYes = False
''''''        Line = 0
''''''        LEFTM = 5
''''''        Print #1, Tab(0); repli("-", 96)
''''''        Print #1, Tab(1); "E.& O.E"
''''''        Print #1, Tab(1); kkk!COURT; Tab(65); "FOR " + Trim(kkk!CNAME)
''''''        Print #1, ""
''''''        Print #1, Tab(1); Chr(27) + Chr(71); "Continued on Page : " & Pno; Chr(27) + Chr(72)
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Print #1, ""
''''''End If
''''''If Printheader = True Then
''''''   If Not kkk.BOF Then
''''''     Print #1, ""
''''''     Print #1, ""
''''''     Print #1, ""
''''''     Print #1, ""
''''''     Print #1, Chr(27) + Chr(71); Chr(27) + Chr(77) + Chr(14)
''''''     Print #1, Tab(((paperWidth - (Len(Trim(kkk!CNAME)) * 2)) / 2)); Chr(27) + Chr(77) + Chr(14); Trim(kkk!CNAME)
''''''     Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(77); dspace(Trim(kkk!add1))
''''''     Print #1, Tab((paperWidth - (Len(Trim(kkk!phone1)) * 2)) / 2); Trim(kkk!phone1) & "," & Trim(kkk!phone2)
''''''     Line = Line + 8
''''''   End If
''''''Else
''''''     Print #1, ""
''''''     Print #1, ""
''''''     Print #1, ""
''''''     Print #1, ""
''''''     Print #1, ""
''''''     Print #1, ""
''''''     Print #1, ""
''''''     Print #1, Chr(27) + Chr(77)
''''''     Line = Line + 8
''''''End If
''''''Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CREDIT NOTE")))) / 2 - 3); Chr(14); "CREDIT NOTE"; Chr(20); Tab(54); IIf(Printheader = True, kkk!uptt, "")
''''''Line = Line + 1
''''''If Printheader = True Then
''''''   Print #1, Tab(63); kkk!cst
''''''   Line = Line + 1
''''''End If
''''''If Printheader = False Then
''''''   Print #1, ""
''''''   Line = Line + 1
''''''End If
''''''Print #1, repli("-", 96)
''''''Line = Line + 1
''''''If rs1.State = 1 Then rs1.close
''''''rs1.Open "CREDITA", CON, adOpenDynamic, adLockReadOnly, adCmdTable
''''''rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
''''''If Not rs1.EOF Then
'''''''    Print #1, Chr(27) + Chr(71); " To,  S.L. Code : "; Tab(20); Mid$(rs1!SUBLEDGER, 1, 5); Tab(46); "C/Note No. : "; Trim(rs1!INVOICENO); Tab(74); "Dated     : "; Chr(27) + Chr(72); rs1!invoicedate
''''''
''''''    Print #1, Chr(27) + Chr(71); " To,: "; Tab(2); Mid$(rs1!agentname, 1, 20); Tab(40); "C/Note No. : "; Trim(rs1!INVOICENO); Tab(74); "Dated     : "; Chr(27) + Chr(72); rs1!invoicedate
''''''    If kkk.State = 1 Then kkk.close
''''''''    kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
''''''''    If Not kkk.EOF Then
''''''''       Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE; Tab(45); Chr(27) + Chr(71); "Bilty No.  : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(75); Chr(27) + Chr(71); "Dated     : "; Chr(27) + Chr(72); IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
''''''
''''''''''      Print #1, Tab(3); IIf(IsNull(kkk!ADDRESS1), " ", kkk!ADDRESS1); Tab(45); Chr(27) + Chr(71); "Freight    : "; Chr(27) + Chr(72); Trim(rs1!freight); Tab(75); Chr(27) + Chr(71); "Bundle(s) : "; Chr(27) + Chr(72); Trim(rs1!bundles)
''''''        Print #1, Chr(27) + Chr(71); "Bilty No. : "; Chr(27) + Chr(71); Trim(rs1!biltyno); Tab(20); Chr(27) + Chr(71); "Bilty Date. : "; Chr(27) + Chr(71); Trim(rs1!biltydate)
''''''        Print #1, Chr(27) + Chr(71); "Freight    : "; Chr(27) + Chr(72); Trim(rs1!freight); Tab(20); Chr(27) + Chr(71); "Demurrage   : "; Chr(27) + Chr(71); Trim(rs1!marka); Tab(40); Chr(27) + Chr(71); "Bundle(s)   : "; Chr(27) + Chr(72); Trim(rs1!bundles); Tab(70); Chr(27) + Chr(71); "Weight(s)  : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."
'''''''''''       Print #1, Tab(3); IIf(IsNull(kkk!ADDRESS2), " ", kkk!ADDRESS2); Tab(45); Chr(27) + Chr(71); "Demurrage  : "; Chr(27) + Chr(72); Trim(rs1!marka); Tab(75); Chr(27) + Chr(71); "Weight    : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."
''''''''       Print #1, Tab(3); IIf(IsNull(kkk!ADDRESS3), " ", kkk!ADDRESS3); Chr(27) + Chr(72); Tab(43); Chr(27) + Chr(71); "Agent Name  : "; Chr(27) + Chr(72); Trim(rs1!agentname)
''''''''       kkk.close
''''''       Print #1, Chr(27) + Chr(71); repli("-", 96)
''''''       Print #1, Tab(0); "S.No."; Tab(15); "Book Description"; Tab(50); "Quantity"; Tab(62); "Rate"; Tab(74); "Amount"; Tab(86); "Net Amount"
''''''       Print #1, repli("-", 96); Chr(27) + Chr(72)
''''''       Line = Line + 8
''''''    'End If
''''''    If called1 Then
''''''        called1 = False
''''''        GoTo printagain1
''''''    End If
''''''    If called2 Then
''''''        called2 = False
''''''        GoTo printagain2
''''''    End If
''''''    If kk.State = 1 Then kk.close
''''''    kk.Open "select * from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " order by printorder,sno ", CON, adOpenDynamic, adLockReadOnly, adCmdText
''''''    If Not kk.BOF Then
''''''        kk.MoveFirst
''''''        Dim cdiscount As Double
''''''        Dim sno As Integer
''''''        Dim tdata As ADODB.Recordset
''''''        Set tdata = New ADODB.Recordset
''''''        sno = 1
''''''        Do While Not kk.EOF
''''''            cdiscount = kk!PRINTORDER
''''''            Do While kk!PRINTORDER = cdiscount
''''''                vdis = kk!discount
''''''                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
''''''                Print #1, Tab(0); rsets(Trim(str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(str(kk!quantity)), 5); Tab(58); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
''''''                totalquantity = totalquantity + kk!quantity
''''''                Line = Line + 1
''''''                If Line > MaxLine - 3 Then
''''''                    called1 = True
''''''                    Pno = Pno + 1
''''''                    FooterYes = True
''''''                    GoTo header
''''''printagain1:
''''''                    called1 = False
''''''               End If
''''''               tdata.close
''''''               If Not kk.EOF Then
''''''                  sno = sno + 1
''''''                  kk.MoveNext
''''''               End If
''''''               If kk.EOF Then
''''''                    Exit Do
''''''               End If
''''''            Loop
''''''            If Line > MaxLine - 5 Then
''''''                    called2 = True
''''''                    Pno = Pno + 1
''''''                    FooterYes = True
''''''                    GoTo header
''''''printagain2:
''''''
''''''                    called2 = False
''''''                End If
''''''                Print #1, Tab(70); repli("-", 12)
''''''                Line = Line + 1
''''''                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " and printorder =" + Trim(str(cdiscount)) + " group by printorder", CON, adOpenDynamic, adLockReadOnly, adCmdText
''''''                If Not tdata.BOF Then
''''''                   Print #1, Tab(68); rsets(Trim(Format(str(tdata(0)), "0.00")), 12)
''''''                   'Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(Str(tdata(0) * cdiscount / 100), "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata(0) - (tdata(0) * cdiscount / 100)), "0.00")), 12)
''''''                   '''Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
''''''                   Print #1, Tab(30); "Less Discount @ " + Trim(Format(str(vdis), "0.00")) + " %"; Tab(68); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(84); rsets(Trim(Format(str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
''''''                   Print #1, Tab(70); repli("-", 12)
''''''                   Line = Line + 3
''''''                   netamount = netamount + tdata!sumamt - tdata!sumdis
''''''                End If
''''''                tdata.close
''''''             Loop
''''''           End If
''''''       End If
''''''       Print #1, repli("-", 96)
''''''       Line = Line + 1
''''''       Print #1, Tab(50); rsets(Trim(str(totalquantity)), 7); Tab(84); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12)
''''''       Line = Line + 1
''''''       If kk.State = 1 Then kk.close
''''''       kk.Open "Select * from CreditC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
''''''       If Not kk.BOF Then
''''''            Do While Not kk.EOF
''''''                If kk!amount > 0 Then
''''''                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
''''''                        netamount = netamount - kk!amount
''''''                    Else
''''''                        netamount = netamount + kk!amount
''''''                    End If
''''''                    If kk!rate > 0 Then
''''''                        Print #1, Tab(60); Trim(kk!Text) + "    " + Trim(Format(str(Round(kk!rate, 2)), "0.00")); Tab(84); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
''''''                    Else
''''''                        Print #1, Tab(60); Trim(kk!Text); Tab(84); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
''''''                    End If
''''''                    Line = Line + 1
''''''                End If
''''''                If Not kk.EOF Then
''''''                    kk.MoveNext
''''''                End If
''''''            Loop
''''''        End If
''''''        Print #1, Tab(84); repli("-", 12)
''''''        Print #1, Chr(27) + Chr(71); Tab(45); "NET AMOUNT Cr. TO YOUR A/C :"; Tab(85); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
''''''        Line = Line + 2
''''''        VNetamt = netamount
''''''        If kk.State = 1 Then kk.close
''''''        kk.Open "Select * from CreditA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenDynamic, adLockReadOnly, adCmdText
''''''        If Not kk.BOF Then
''''''            If kk!txt1a <> 0 Then
''''''                Print #1, Tab(60); kk!txt1; Tab(84); rsets(Trim(Format(str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
''''''                Line = Line + 1
''''''                netamount = netamount + Round(kk!txt1a, 2)
''''''             End If
''''''             If kk!txt2a <> 0 Then
''''''                 Print #1, Tab(60); kk!txt2; Tab(84); rsets(Trim(Format(str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
''''''                 Line = Line + 1
''''''                 netamount = netamount + Round(kk!txt2a, 2)
''''''             End If
''''''             If kk!baa <> 0 Then
''''''                 Print #1, Tab(60); "BY BANK "; Tab(84); rsets(Trim(Format(str(Abs(Round(kk!baa, 2))), "0.00")), 12)
''''''                 Line = Line + 1
''''''                 netamount = netamount - Round(kk!baa, 2)
''''''             End If
''''''        End If
''''''        Print #1, Tab(84); repli("-", 12)
''''''        Line = Line + 1
''''''      ' PRINT THE FOOTER IN INVOICE START
''''''        Do While Line < 61
''''''            Print #1, ""
''''''            Line = Line + 1
''''''        Loop
''''''        Print #1, Tab(0); Chr(27) + Chr(71); toword(Round(VNetamt, 2)); Chr(27) + Chr(72)
''''''        Print #1, Tab(0); repli("-", 96)
''''''        Dim tempdata As ADODB.Recordset
''''''        Set tempdata = New ADODB.Recordset
''''''        CNSetup
''''''        tempdata.Open "setup1", CON, adOpenDynamic, adLockReadOnly, adCmdTable
''''''        Print #1, Tab(1); "E.& O.E"
''''''        Print #1, Tab(1); tempdata!COURT; Tab(65); "FOR " + Trim(tempdata!CNAME)
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Print #1, ""
''''''        Close #1
''''''        PrintOption.Show
End Sub
  
Sub CREDITCalc()
    'OTHERcredit.calc
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     mna.Caption = Format(Round((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
End Sub
Sub CREDITAbandon()
        Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
        Command1.Enabled = True
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If kk.State = 1 Then
   kk.Close
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
                CRITNOTE.customercode.Enabled = False
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
        Unload OTHERCredit
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
                            rs.Close
                            templost = False
                            Exit Function
                        Else
                            Grid1.Text = rs(0)
                            Grid1.col = 2
                            Grid1.Text = rs(1)
                         'If Not edit Then
                                Grid1.col = 3
                                If Trim(Grid1.Text) = "" Then
                                    Grid1.Text = 0
                                End If
                                q = Val(Grid1.Text)
                                
                                
                                Grid1.col = 4
                                If Grid1.Text = "" Then
                                Grid1.Text = Format(rs(4), "0.00")
                                End If

                                Grid1.col = 5
                                If Grid1.Text = "" Then
                                Grid1.Text = Format(rs(3), "0.00")            'rs(3)
                                End If
                                r = rs(3)
                                        
                                Grid1.col = 6
                                If Grid1.Text = "" Then
                                   Grid1.Text = Format(rs(4), "0.00")
                                End If
                                        D = rs(4)
                             '        End If
                                    Grid1.col = 7
                                    Grid1.Text = Format(Round(q * r, 2), "0.00")
                                    Grid1.col = 8
                                    Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
                            'Else
                               If Grid1.Text = "" And addmode = False Then
                                     If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        kk.Close
                                        Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
                                        Grid1.col = 4
'''                                        If kk.BOF Then
'''                                             GoTo abc
'''                                        End If
                                        Grid1.Text = Format(kk(0), "0.00")
                                        Grid1.col = 6
                                        Grid1.Text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = rs(3)
                            
                                    End If
                            
                                End If
                            
                            'End If
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
        CREDITCalc
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
                        
                           Commandother.SetFocus
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
''''''                                Set kk = CON.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
''''''                                Grid1.col = 6
''''''                                If Trim(kk(0)) <> "" Then
''''''                                    tempstr = Trim(kk(0))
''''''                                    kk.Close
''''''                                    Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(rs(2)) + "'")
''''''                                    Grid1.col = 4
''''''                                    If kk.BOF Then
''''''                                        GoTo abc
''''''                                    End If
''''''                                    Grid1.Text = Format(kk(0), "0.00")
''''''                                    Grid1.col = 6
''''''                                    Grid1.Text = Format(kk(0), "0.00")
''''''                                    D = kk(0)
''''''                                Else
''''''abc:
                                    Grid1.col = 4
                                    Grid1.Text = Format(rs(4), "0.00")
                                    Grid1.col = 6
                                    Grid1.Text = Format(rs(4), "0.00")
                                    D = rs(4)
                                'End If
                                Grid1.col = 7
                                Grid1.Text = Round(q * r, 2)
                                Grid1.col = 8
                                Grid1.Text = Round((q * r) * (D / 100), 2)
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
        For I = 1 To maxrow
            Grid1.row = I
            Grid1.col = 7
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
        Next
        CREDITCalc
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


Private Sub cmbAgentName_LostFocus()
If cmbAgentName.Text = "" Then
   MsgBox "Enter a Agent Name.. "
   'cmbAgentName.SetFocus
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

Private Sub Command1_Click()
Printheader = False
printinvoice
End Sub

Private Sub Command1_LostFocus()
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

End Sub

Private Sub Commandabandon_Click()
CREDITAbandon
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

End Sub
Private Sub Commandadd_Click()
On Error Resume Next
    addmode = True
    Edit = False
    CREDITAbandon
    Dim rs As ADODB.Recordset
    addoredit = True
    addmode = True
    Set rs = New ADODB.Recordset
    Dim TEMPNUM As Integer
    If Edit = False Then
       'If CON.Execute("Select max(invoiceno) from CREDITA")(0) >= Val(Trim(Me.I_NO.Text)) Then
              Me.I_NO.Text = CON.Execute("Select max(invoiceno) from CREDITA")(0) + 1
              rs.Open "tempCRITNOTE", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
              If rs.BOF Then
                 rs.AddNew
              End If
               Me.I_NO.Text = rs!In + 1
               rs!In = Val(Me.I_NO.Text)
               rs.update
               rs.Close
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
    
    cboGodown.ListIndex = 0
    
    
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
    rs.Close
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

CREDITCalc

End Sub

Private Sub Commanddelete_Click()


'==================

    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select * from credita where INVOICENO=" & Trim(I_NO.Text) & "", CON
    If rs1.EOF = False Then
       
    If rs_h.State = 1 Then rs_h.Close
    rs_h.Open "select * from credita where INVOICENO=" & Trim(I_NO.Text) & "", CON
    'If rs_h.Fields("Print_yes").Value = "y" Then
       If rs1!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
       End If
    
    'End If
    
    End If
    
'=====================




If rs_v.State = 1 Then rs_v.Close
rs_v.Open "select BDelete,Bedit,BSave from setup where UserId=" & UId & "", CONINFO
If rs_v.EOF = False Then
If rs_v!bDelete = False Then
   MsgBox "You Can'nt Delete ...", vbCritical
   Exit Sub
End If
End If



If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
                CON.Execute ("delete * from CREDITA where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete * from CREDITB where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("delete * from CREDITC where INVOICENO = " + Trim(I_NO.Text))
                CREDITAbandon
End If
End Sub

Private Sub Commandedit_Click()
   
   
   
   If rs_v.State = 1 Then rs_v.Close
rs_v.Open "select BDelete,Bedit,BSave from setup where UserId=" & UId & "", CONINFO
If rs_v.EOF = False Then
If rs_v!bedit = False Then
   MsgBox "You Can'nt Edit ...", vbCritical
   Exit Sub
End If
End If


   
   
   
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
    CON.Execute ("delete * from CREDITCtmp where INVOICENO = " & CRITNOTE.I_NO & "")
    CON.Execute ("insert into CREDITCtmp  select * from CREDITC where INVOICENO = " + Trim(I_NO.Text))
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
    OTHERCredit.Top = 0
    OTHERCredit.Left = 0
    OTHERCredit.Visible = False
    DoEvents
   
    
    
End Sub
Private Sub Commandother_Click()
    Me.Enabled = False
   
    OTHERCredit.Top = 0
    OTHERCredit.Left = 0
    'Unload OTHERCredit
    'Load OTHERcredit
    OTHERCredit.Show
   
End Sub
Private Sub CommandPrint_Click()
Printheader = True
printinvoice

End Sub

Private Sub CommandPrint_LostFocus()
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

End Sub

Private Sub Commandreturn_Click()
   Dim rs As New ADODB.Recordset
   Commandsave.Enabled = True
   rs.Open "tempCRITNOTE", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
   If rs.BOF Then
       rs.AddNew
   End If
   rs!In = CON.Execute("Select max(invoiceno) from CREDITA")(0)
   rs.update
   rs.Close
   Unload Me
   addoredit = False
    
End Sub
Private Sub Commandsave_Click()
    
    
    
    
    
    
'==================

    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select * from credita where INVOICENO=" & Trim(I_NO.Text) & "", CON
    If rs1.EOF = False Then
       
    If rs_h.State = 1 Then rs_h.Close
    rs_h.Open "select * from credita where INVOICENO=" & Trim(I_NO.Text) & "", CON
    'If rs_h.Fields("Print_yes").Value = "y" Then
       If rs1!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
       End If
    
    'End If
    
    End If
    
'=====================
    
    
    
    
       
    Dim SAVED As Boolean
    Dim LAMOUNT As Double
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    
    If rs_v.State = 1 Then rs_v.Close
rs_v.Open "select BDelete,Bedit,BSave from setup where UserId=" & UId & "", CONINFO
If rs_v.EOF = False Then
If rs_v!bsave = False Then
   MsgBox "You Can'nt Save ...", vbCritical
   Exit Sub
End If
End If


    
    
    
     If Edit = False And addmode = False Then
      Me.Commandsave.Enabled = False
      Exit Sub
    End If
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
    'If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(customercode.Text) <> "" Then
        If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(Me.cmbAgentName.Text) <> "" Then
        If Edit Then
            CON.Execute ("delete * from CREDITA where INVOICENO = " + Trim(I_NO.Text))
            CON.Execute ("delete * from CREDITB where INVOICENO = " + Trim(I_NO.Text))
            CON.Execute ("delete * from CREDITC where INVOICENO = " + Trim(I_NO.Text))
        End If
        If rs.State = 1 Then
            rs.Close
        End If
            LAMOUNT = 0
            rs.Open "select * from CREDITA  where invoiceno<=0", CON, adOpenDynamic, adLockOptimistic
            If Not Edit Then
again:
               If CON.Execute("Select max(invoiceno) from CREDITA")(0) >= Val(Trim(Me.I_NO.Text)) Then
                    'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'GoTo again
               End If
            End If
            rs.AddNew
            rs!INVOICENO = Val(Me.I_NO.Text)
            rs!godown = cboGodown.Text
            rs!invoicedate = Me.i_dt.Text
            'rs!Genledger = Trim(Me.Genledger.Text)
            'rs!SUBLEDGER = Trim(Me.customercode.Text)
            rs!marka = Trim(Me.marka.Text)
            rs!bundles = Trim(Me.bundles)
            rs!biltyno = Trim(Me.biltno.Text)
            If Trim(Me.bdated) = Trim("__/__/____") Then
                rs!biltydate = Null '"__/__/____"
            Else
                rs!biltydate = Me.bdated & ""
            End If

            rs!freight = Me.freight & ""
            rs!weight = Me.weight & ""
            rs!netamount = Round(Val(Trim(Me.mna.Caption)), 2)
            rs!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
            rs!txt1 = Trim(OTHERCredit.T1TEXT.Text)
            rs!txt1a = Val(Trim(OTHERCredit.T1.Text))
            rs!txt2 = Trim(OTHERCredit.T2TEXT.Text)
            rs!txt2a = Val(Trim(OTHERCredit.T2.Text))
            rs!baa = Val(Trim(OTHERCredit.T3TEXT.Text))
            rs!agentname = cmbAgentName.Text
            Dim trs As New ADODB.Recordset
            trs.Open " SELECT DISTCODE  FROM SLEDGER  WHERE SUBLEDGER='" & customercode.Text & "'", CON, adOpenStatic, adLockOptimistic, adCmdText
            If Not trs.BOF Then
               rs!District = trs!distcode & ""
            Else
               rs!District = ""
            End If
            trs.Close
err1:
            If Not Edit Then
                If CON.Execute("Select max(invoiceno) from CREDITA")(0) >= Val(Trim(Me.I_NO.Text)) Then
                    'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'rs!INVOICENO = Val(Me.I_NO.Text)
                    On Error GoTo err1
                End If
            End If
            rs.update
            On Error GoTo 0
            rs.Close
            rs.Open "select * from CREDITB where invoiceno<=0", CON, adOpenDynamic, adLockOptimistic
            Dim I As Integer
            RRRR = Grid1.row
            CCCC = Grid1.col
            For I = 1 To maxrow
                Grid1.row = I
                Grid1.col = 1
                If Trim(Grid1.Text) <> "" Then
                    Grid1.col = 3
                    If Val(Trim(Grid1.Text)) > 0 Then
                        rs.AddNew
                        Grid1.col = 1
                        rs!INVOICENO = Val(Me.I_NO.Text)
                        rs!invoicedate = Me.i_dt.Text
                       ' rs!Genledger = Trim(Me.Genledger.Text)
                       ' rs!SUBLEDGER = Trim(Me.customercode.Text)
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
                        rs.update
                    End If
                End If
            Next
            rs.Close
            Grid1.TopRow = 1
            
            rs.Open "select * from CREDITC where invoiceno<=0", CON, adOpenDynamic
            '/******
            Dim temprs As ADODB.Recordset
            Set temprs = New ADODB.Recordset
            For I = 1 To OTHERCredit.mrow
                    OTHERCredit.Grid1.row = I
                    OTHERCredit.Grid1.col = 0
                    If Trim(OTHERCredit.Grid1.Text) <> "" Then
                        rs.AddNew
                        rs!INVOICENO = Val(Me.I_NO.Text)
                        rs!invoicedate = Me.i_dt.Text
                        rs!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
                        rs!Text = Trim(OTHERCredit.Grid1.Text)
                        If temprs.State = 1 Then
                            temprs.Close
                        End If
                        If Edit Then
                        temprs.Open "select * from CREDITCtmp WHERE INVOICENO=" & CRITNOTE.I_NO & "", CON, adOpenStatic, adLockReadOnly, adCmdText
                        If OTHERCredit.Grid1.Text <> "" Then
                                temprs.Find "TEXT='" + Trim(OTHERCredit.Grid1.Text) + "'"
                        '        rs!Genledger = temprs!Genledger & ""
                         '       rs!SUBLEDGER = temprs!SUBLEDGER & ""
                                rs!DebitorCredit = temprs!DebitorCredit & ""
                                rs!RYN = temprs!RYN & ""
                        End If
                        temprs.Close
                        Else
                        temprs.Open "select * from CREDITEND", CON, adOpenStatic, adLockReadOnly, adCmdText
                        If OTHERCredit.Grid1.Text <> "" Then
                           temprs.Find "TEXT='" + Trim(OTHERCredit.Grid1.Text) + "'"
                        '   rs!Genledger = temprs!Genledger & ""
                        '   rs!SUBLEDGER = temprs!SUBLEDGER & ""
                           rs!DebitorCredit = temprs!DebitorCredit & ""
                           rs!RYN = temprs!RYN & ""
                        End If
                        temprs.Close
                        End If
                        OTHERCredit.Grid1.col = 1
                        rs!rate = Val(Trim(OTHERCredit.Grid1.Text))
                        If Val(Trim(OTHERCredit.Grid1.Text)) > 0 Then
                            rs!amount = Round((Me.totalamount - Me.totaldiscount), 2) * Round((Val(Trim(OTHERCredit.Grid1.Text)) / 100), 2)
                        Else
                        OTHERCredit.Grid1.col = 2
                            rs!amount = Val(Trim(OTHERCredit.Grid1.Text))
                        End If
                    rs.update
                    End If
            Next
            rs.Close
            CON.Execute ("delete * from CreditCTmp where INVOICENO = " + Trim(I_NO.Text))
            
          If addmode = True Then

                    rs.Open "tempCRITNOTE", CON1, adOpenStatic, adLockOptimistic, adCmdTable
                    If rs.BOF Then
                        rs.AddNew
                    End If
                    rs!In = CON.Execute("Select max(invoiceno) from CREDITA")(0)
                    rs.update
                    rs.Close
          End If
                    
'''                rs.Open "tempCRITNOTE", CON1, AdopenStatic, adLockOptimistic, adCmdTable
'''                If rs.BOF Then
'''                   rs.AddNew
'''                End If
'''                rs!In = Val(Me.I_NO.Text)
'''                rs.Update
'''                rs.Close
'''                End If
               SAVED = True
        End If
        If SAVED Then
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
        Command1.Enabled = True
        End If
        addmode = False
        addoredit = False
        SetButton Commandadd, Commandedit, Commandsave, Commanddelete
        Me.Commandsave.Enabled = False
   End Sub

Private Sub Commandsearch_Click()
    Me.Enabled = False
    Call searchscreen.tempr(13, "CREDITITEMNOTE")
End Sub

Private Sub customercode_LostFocus()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select * from sledger where gledger='" + Genledger.Text + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
    rs.Find "subledger='" + Trim(customercode.Text) + "'"
    If rs.EOF Then
        customercode.SetFocus
        HIT
        rs.Close
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
    If rs.State = 1 Then rs.Close
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
        Dim trs3 As New ADODB.Recordset
        trs3.Open "Select * from cnf1a  where cnn = " & Val(I_NO.Text) & "", CON, adOpenStatic, adLockReadOnly, adCmdText
        If trs3.RecordCount > 0 Then
                MsgBox "CREDIT Note File already exist..."
                I_NO.SetFocus
                HIT
                Exit Sub
        End If
   
        
        
        
        
           If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("grid1")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("tempmeb")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bookname")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("weight")) Then
              SendKeys ("{TAB}")
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
On Error Resume Next


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
    Grid1.Left = 90

    Grid1.Top = 1470
   ' Set CON = New ADODB.Connection
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
    Grid1.ColWidth(0) = 300
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 2200
    Grid1.ColWidth(3) = 750
    Grid1.ColWidth(4) = 750
    Grid1.ColWidth(5) = 900
    Grid1.ColWidth(6) = 900
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Me.CommandPrint.Enabled = True
    rs.Open "select * from books", CON, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.Bookcode.AddItem rs(0)
            Me.Bookname.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    Genledger.AddItem "SUNDRY DEBTORS"
    Genledger.Text = "SUNDRY DEBTORS"
    rs.Open "select SUBLEDGER from sledger where gledger='" + Trim(Genledger.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            Me.customercode.AddItem rs(0)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    

     '*******Agent  combo fill
    rs.Open "select Distinct Agentname from AgentMaster order by agentname", CON, adOpenStatic, adLockReadOnly, adCmdText
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
 
    
    rs.Open "select Godwn from godownMaster order by id"
    While rs.EOF = False
          cboGodown.AddItem rs(0)
          rs.MoveNext
    Wend
    rs.Close

    
    
    Bookname.Height = 1935
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
    rs.Open "TEMPCRITNOTE", CON1, adOpenStatic, adLockOptimistic, adCmdTable
    If Not rs.BOF Then
       Me.I_NO.Text = rs!In
       CRITNOTE.Enabled = True
       CRITNOTE.Edit = False
       trs.Open "Select * from cnf1a  where cnn = " & Val(I_NO.Text) & "", CON, adOpenStatic, adLockReadOnly, adCmdText
       If trs.RecordCount > 0 Then
                MsgBox "CREDIT Note File already exist..."
                'I_NO.SetFocus
                'HIT
                Exit Sub
       End If
       CRITNOTE.I_NO_LostFocus
       CRITNOTE.I_NO.Enabled = False
       lastrow = 0
       lastcol = 1
       Dim ctl As Control
       For Each ctl In CRITNOTE.Controls
            If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = False
            End If
            If UCase(Trim(ctl.Name)) = UCase(Trim(CRITNOTE.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(CRITNOTE.Commandall.Name)) Then
                ctl.Enabled = False
            End If
       Next
       CRITNOTE.Picture5.Enabled = True
       addoredit = False
       SendKeys "{TAB}"
    Else
       kk.Open "SELECT MAX(INVOICENO) FROM CREDITA", CON, adOpenStatic, adLockReadOnly, adCmdText
       If kk(0) <> "" Then
            Me.I_NO.Text = Trim(str(kk(0) + 1))
       Else
            Me.I_NO.Text = "1"
       End If
        kk.Close
   End If
   rs.Close
   Commanddelete.Enabled = True
   Commandedit.Enabled = True
   Commandsave.Enabled = False
   Command1.Enabled = True
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
''    If Trim(Me.customercode.Text) <> "" Then
''        If Me.customercode.Enabled = True Then
''            Me.customercode.Enabled = False
''        End If
        
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
       i_dt.Enabled = True
        i_dt.SetFocus
    End If
    Dim tRS1 As New ADODB.Recordset
    Dim trs2 As New ADODB.Recordset
    
    If trs2.State = 1 Then trs2.Close
    trs2.Open "Select invoiceno as cn from credita", CON, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount <= 0 Then
       Exit Sub
    Else
        If tRS1.State = 1 Then tRS1.Close
        tRS1.Open "Select min(invoiceno) as mid,invoicedate from credita group by invoiceno,invoiceDate", CON, adOpenDynamic, adLockOptimistic
        If tRS1.RecordCount > 0 Then
             
            If CDate(i_dt) <= tRS1!invoicedate Then
               If CDate(i_dt) <> tRS1!invoicedate Then
                 MsgBox "Please select valid Credit  No. for this date.."
                 I_NO.SetFocus
                 Exit Sub
               End If
            End If
        End If
    End If
    
    
    If trs2.State = 1 Then trs2.Close
    trs2.Open "Select max(invoiceno) as mid from credita where  invoicedate <= cdate('" & i_dt.Text & "')-1", CON, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount > 0 Then
        If IsNull(trs2!Mid) = False Then
            If Val(I_NO.Text) >= trs2!Mid Then
               If tRS1.State = 1 Then tRS1.Close
               tRS1.Open "Select  min(InvoiceNo)as m2 from credita where invoicedate >= cdate('" & i_dt.Text & "')+1", CON, adOpenDynamic, adLockOptimistic
               If tRS1.RecordCount > 0 Then
                  If IsNull(tRS1!m2) <> True Then
                     If Val(I_NO.Text) <= tRS1!m2 Then
                       
                     Else
                         MsgBox "Please select valid Credit No. for this date.."
                         I_NO.SetFocus
                     End If
                  End If
               End If
            
            Else
               MsgBox "Please select valid Credit  No. for this date.."
               I_NO.SetFocus
            End If
            
        Else
        'If i_dt.Enabled = True Then
         '  MsgBox "Please select valid Credit  No for this date.."
          ' I_NO.SetFocus
        'End If
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
Dim rs As ADODB.Recordset
Dim rs3  As ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs3 = New ADODB.Recordset


If Val(inviceNo) > 0 Then
I_NO.Text = inviceNo
  ' cmdButto
End If

inviceNo = ""



If Trim(I_NO.Text) = "" Then
        MsgBox "Credit Note no cannot be null"
        I_NO.SetFocus
Else
    If rs.State = 1 Then rs.Close
    
    'rs.Open "CREDITA", con, AdopenStatic, adLockReadOnly, adCmdTable
    rs.Open "Select * from  CREDITA where INVOICENO = " + Trim(I_NO.Text) + "", CON, adOpenStatic, adLockReadOnly
    rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
    rs3.Open "Select * from cnf1a  where cnn = " & Val(I_NO.Text) & "", CON, adOpenStatic, adLockReadOnly, adCmdText
    If rs3.RecordCount > 0 Then
                MsgBox "CREDIT Note File already exist..."
                I_NO.SetFocus
                HIT
                Exit Sub
   End If
   
   
    If rs.EOF Then
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
   Dim ctl As Control
   For Each ctl In Me.Controls
       If Not TypeOf ctl Is CommandButton Then
          ctl.Enabled = True
       End If
   Next
        Me.Commandother.Enabled = True
        I_NO.Text = rs!INVOICENO
        cboGodown.Text = rs!godown & ""
        Me.i_dt.Text = rs!invoicedate
        Me.Genledger.Text = Trim(rs!Genledger)
        Me.customercode.Text = Trim(rs!SUBLEDGER)
        Me.textbox.Text = Trim(rs!SUBLEDGER)
        Me.marka.Text = Trim(IIf(IsNull(rs!marka), "", rs!marka))
        Me.bundles = Trim(rs!bundles)
        Me.biltno.Text = Trim(rs!biltyno)
        
        If IsNull(rs!biltydate) Then
           Me.bdated = "__/__/____"
        Else
           Me.bdated = rs!biltydate
        End If
        Me.freight = Trim(rs!freight)
        Me.weight = Trim(rs!weight)
        Me.labelbybank = Trim(rs!baa)
        mna.Caption = rs!netamount
        Me.cmbAgentName.Text = IIf(IsNull(rs!agentname), "", rs!agentname)
        rs.Close
'*/**/*/*/*/*//*/*
        If rs.State = 1 Then
            rs.Close
        End If
 rs.Open "Select * from CREDITB where INVOICENO =" + Trim(I_NO.Text) + " order by SNO ", CON, adOpenStatic, adLockReadOnly
       '' rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
        Grid1.TopRow = 1
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
       ' templost = True
    End If
    
    
     Me.Commandother.Enabled = True






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
                    rs.Open "books", CON, adOpenStatic, adLockReadOnly, adCmdTable
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

Private Sub weight_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            Me.Grid1.SetFocus
            Grid1.col = 1
            Grid1.row = 1
            Grid1_Click
End If



''''If KeyAscii = 13 Then
''''        If Trim(Me.customercode.Text) <> "" Then
''''            Grid1.col = 1
''''            Grid1.row = 1
''''            Grid1_Click
''''        Else
''''             Me.textbox.SetFocus
''''            'Me.customercode.SetFocus
''''        End If
''''    End If
   
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
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = True
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
        kkk.Close
    End If
    CNSetup
    kkk.Open "select * from setup1", CON, adOpenStatic, adLockReadOnly, adCmdText
    If flagyes = True Then

    
      If Not kkk.BOF Then
        Print #1, Chr(27) + Chr(15) + Chr(14)
        Print #1, Tab(T1); Chr(27) + Chr(15) + Chr(14); Trim(kkk!CNAME) '
        Print #1, Tab(T2 - 7); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
        Print #1, Tab(T3); Trim(kkk!phone1)
        Line = Line + 4
    End If
  
    Print #1, repli("-", 150)

  End If
    If rs1.State = 1 Then
        rs1.Close
    End If
    rs1.Open "CREDITA", CON, adOpenStatic, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
        Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!SUBLEDGER; Tab(T5); "Credit Note No. : "; Trim(rs1!INVOICENO); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!invoicedate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.Close
            End If
            kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE
                'Print #1, Tab(3); kkk!ADDRESS1; Tab(t5); "Order by : "; Trim(rs1!ORDERBY); Tab(t8 + 5); "Dt. "; Tab(t8 + 12); rs1!ORDERDATE
                Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.: "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!biltydate
                kkk.Close
                Print #1,
                'Print #1, "Through  :"; Tab(12); Trim(rs1!through) + ", " + Trim(rs1!through1)
                'Print #1, "Station  :"; Tab(12); Trim(rs1!station); Tab(t5); "Demurrage.        : "; Trim(rs1!marka)
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
                kk.Close
            End If
            kk.Open "select * from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenStatic, adLockReadOnly, adCmdText
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
                        tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
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
                        tdata.Open "select sum(amount) from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenStatic, adLockReadOnly, adCmdText
                        If Not tdata.BOF Then
                            Print #1, Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0), 2)), "0.00")), 12)
                            Print #1, Tab(T5); "Less Discount @ " + Trim(Format(str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
                            netamount = netamount + Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)
                        End If
                        tdata.Close
                        Print #1, Tab(T7); repli("-", 22)
                Loop
            End If
           End If
           Print #1, Tab(T5 - 4); rsets(Trim(str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12)
           Print #1, Tab(T6); repli("-", 22)
           If kk.State = 1 Then
                kk.Close
           End If
           kk.Open "Select * from CREDITC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
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
           kk.Close
           kk.Open "Select * from CREDITA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
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
            Print #1, Tab(0); repli("-", 120)
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            Dim LEFTM As Integer
            LEFTM = 5
            CNSetup
            tempdata.Open "setup1", CON, adOpenStatic, adLockReadOnly, adCmdTable
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

Sub Backupprintinvoice()
Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = True
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = True
Me.Commandabandon.Enabled = True
Me.CommandPrint.Enabled = True
Me.Command1.Enabled = True
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
      kkk.Close
End If
 CNSetup
 kkk.Open "select * from setup1", CON, adOpenStatic, adLockReadOnly, adCmdText
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
   rs1.Close
End If





If rs1.State = 1 Then
   rs1.Close
End If
'Print #1, Chr(27) + Chr(14)
'line = line + 1
If rs1.State = 1 Then
    rs1.Close
End If
rs1.Open "CreditA", CON, adOpenStatic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To, S.L. Code : "; Tab(T1 + 8); Mid$(rs1!SUBLEDGER, 1, 5); Tab(T5); "C/Note No. : "; Trim(rs1!INVOICENO); Tab(T8); "Dated     : "; rs1!invoicedate
    kine = libe + 1
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from sledger where subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE; Tab(T5); "Bilty No.  : "; Trim(rs1!biltyno); Tab(T8); "Dated     : "; IIf(IsNull(rs1!ORDERDATE), "  /  /    ", rs1!ORDERDATE)
        Print #1, Tab(3); kkk!ADDRESS1; Tab(T5); "Freight    : "; Trim(rs1!freight); Tab(T8); "Bundle(s) :  "; Trim(rs1!bundles);
        Print #1, Tab(3); kkk!ADDRESS2; Tab(T5); "Demurrage  : "; Trim(rs1!marka); Tab(T8); "Weight    :  "; Trim(rs1!weight)
        Print #1, Tab(3); kkk!ADDRESS3; Chr(27) + Chr(72)
        kkk.Close
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
    kk.Open "select * from CreditB where invoiceno=" + Trim(rs1!INVOICENO) + " order by printorder", CON, adOpenStatic, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
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
                tdata.Open "select sum(amount) from CreditB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenStatic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                    Print #1, Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0), 2)), "0.00")), 12)
                    Print #1, Tab(T5); "Less Discount @ " + Trim(Format(str(Round(cdiscount, 2)), "0.00")) + " %"; Tab(T7 - 1); rsets(Trim(Format(str(Round(tdata(0) * cdiscount / 100, 2)), "0.00")), 12); Tab(T8 + 5); rsets(Trim(Format(str(Round(tdata(0) - (tdata(0) * cdiscount / 100), 2)), "0.00")), 12)
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
        Print #1, repli("-", 150)
        Print #1, Tab(T5 - 4); rsets(Trim(str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12)
       'by vk Print #1, Tab(t8); repli("-", 22)
        Line = Line + 2
        If kk.State = 1 Then
             kk.Close
        End If
        kk.Open "Select * from CreditC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5 + 10); Trim(kk!Text) + "    " + Trim(Format(str(Round(kk!rate, 2)), "0.00")); Tab(T8 + 5); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5 + 10); Trim(kk!Text); Tab(T8 + 5); rsets(Trim(Format(str(Round(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T8); repli("-", 22)
            Print #1, Chr(27) + Chr(71); Tab(T5 - 10); "NET AMOUNT Cr. TO YOUR A/C :"; Tab(T8 + 6); rsets(Trim(Format(str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
            Line = Line + 2
            VNetamt = netamount
        End If
        kk.Close
        kk.Open "Select * from CreditA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(T5 + 10); kk!txt1; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + Round(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(T5 + 10); kk!txt2; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + Round(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(T5 + 10); "BY BANK "; Tab(T8 + 5); rsets(Trim(Format(str(Abs(Round(kk!baa, 2))), "0.00")), 12)
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
        tempdata.Open "setup1", CON, adOpenStatic, adLockReadOnly, adCmdTable
        Print #1, Tab(1); "E.& O.E"
        Print #1, Tab(0); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!CNAME)) * 0.75))); "FOR " + Trim(tempdata!CNAME)
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

