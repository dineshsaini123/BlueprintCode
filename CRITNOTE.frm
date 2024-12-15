VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CRITNOTE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6285
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CRITNOTE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbAgentName 
      Height          =   315
      Left            =   5880
      TabIndex        =   7
      Top             =   360
      Width           =   3795
   End
   Begin VB.CommandButton Commandall 
      Caption         =   "All Books"
      Height          =   375
      Left            =   1020
      TabIndex        =   46
      Top             =   5160
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Commandother 
      Caption         =   "&Other"
      Height          =   375
      Left            =   210
      TabIndex        =   22
      Top             =   5190
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2955
      Left            =   60
      TabIndex        =   39
      Top             =   1440
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   5212
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
      TabIndex        =   34
      Top             =   2550
      Width           =   2295
   End
   Begin VB.ComboBox Bookcode 
      Height          =   1935
      Left            =   540
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   33
      Top             =   2040
      Width           =   2355
   End
   Begin VB.PictureBox Picture5 
      Height          =   465
      Left            =   -30
      ScaleHeight     =   405
      ScaleWidth      =   9495
      TabIndex        =   12
      Top             =   5610
      Width           =   9555
      Begin VB.CommandButton Command1 
         Caption         =   "N&HPrint"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6711
         TabIndex        =   49
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1113
         TabIndex        =   0
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandReturn 
         Caption         =   "&Return"
         Height          =   375
         Left            =   8580
         TabIndex        =   19
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton CommandPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7644
         TabIndex        =   18
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5778
         TabIndex        =   17
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commanddelete 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4845
         TabIndex        =   16
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandabandon 
         Caption         =   "Aba&ndon"
         Height          =   375
         Left            =   3912
         TabIndex        =   15
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandsave 
         Caption         =   "Sa&ve"
         Height          =   375
         Left            =   2970
         TabIndex        =   24
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2046
         TabIndex        =   14
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton Commandhelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   800
      End
   End
   Begin MSMask.MaskEdBox weight 
      Height          =   285
      Left            =   3870
      TabIndex        =   11
      Top             =   1050
      Width           =   1905
      _ExtentX        =   3360
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
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bdated 
      Height          =   315
      Left            =   3510
      TabIndex        =   4
      Top             =   330
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
      Left            =   2310
      TabIndex        =   3
      Top             =   330
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bundles 
      Height          =   285
      Left            =   2340
      TabIndex        =   10
      Top             =   1050
      Width           =   2265
      _ExtentX        =   3995
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
      Left            =   1230
      TabIndex        =   2
      Top             =   330
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
      TabIndex        =   35
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
      TabIndex        =   36
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
      TabIndex        =   37
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
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.ComboBox customercode 
      Height          =   1740
      Left            =   5880
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.ComboBox Genledger 
      Height          =   315
      Left            =   5490
      Sorted          =   -1  'True
      TabIndex        =   38
      Top             =   4470
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox marka 
      Height          =   285
      Left            =   1230
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
      Left            =   5880
      TabIndex        =   5
      Top             =   0
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agent :"
      Height          =   315
      Left            =   4770
      TabIndex        =   50
      Top             =   330
      Width           =   1020
   End
   Begin VB.Label labelbybanklbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By Bank : "
      Height          =   255
      Left            =   2340
      TabIndex        =   48
      Top             =   5250
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label labelbybank 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   47
      Top             =   5280
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   3510
      TabIndex        =   23
      Top             =   30
      Width           =   1200
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dem."
      Height          =   285
      Left            =   1230
      TabIndex        =   27
      Top             =   750
      Width           =   1065
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Quantity : "
      Height          =   255
      Left            =   3450
      TabIndex        =   45
      Top             =   4710
      Width           =   1470
   End
   Begin VB.Label tqu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3420
      TabIndex        =   44
      Top             =   4980
      Width           =   1500
   End
   Begin VB.Label mgd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8190
      TabIndex        =   42
      Top             =   5040
      Width           =   1290
   End
   Begin VB.Label mna 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8190
      TabIndex        =   41
      Top             =   5310
      Width           =   1290
   End
   Begin VB.Label mga 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6900
      TabIndex        =   40
      Top             =   5040
      Width           =   1260
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dated : "
      Height          =   285
      Left            =   1230
      TabIndex        =   32
      Top             =   15
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Credit Note No. : "
      Height          =   285
      Left            =   60
      TabIndex        =   31
      Top             =   15
      Width           =   1155
   End
   Begin VB.Label label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cust Code : "
      Height          =   285
      Left            =   4740
      TabIndex        =   30
      Top             =   30
      Width           =   1065
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Net Amount : "
      Height          =   255
      Left            =   6900
      TabIndex        =   29
      Top             =   5310
      Width           =   1260
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Gross Amount : "
      Height          =   255
      Left            =   6900
      TabIndex        =   28
      Top             =   4740
      Width           =   1260
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bundle(s) : "
      Height          =   285
      Left            =   2340
      TabIndex        =   26
      Top             =   750
      Width           =   1515
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bilty No. : "
      Height          =   285
      Left            =   2310
      TabIndex        =   25
      Top             =   15
      Width           =   1185
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Freight : "
      Height          =   285
      Left            =   60
      TabIndex        =   21
      Top             =   750
      Width           =   1155
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Weight : "
      Height          =   285
      Left            =   3870
      TabIndex        =   20
      Top             =   750
      Width           =   1920
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Total Discount : "
      Height          =   255
      Left            =   8190
      TabIndex        =   43
      Top             =   4740
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
Attribute VB_Name = "CRITNOTE"
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
Dim printheader As Boolean
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
Open "" + App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
FooterYes = False
header:
If kkk.State = 1 Then kkk.Close
    CNSetup
    kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
     Print #1, ""
     Print #1, Chr(27) + Chr(77)
     Line = Line + 8
End If
Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CREDIT NOTE")))) / 2 - 3); Chr(14); "CREDIT NOTE"; Chr(20); Tab(54); IIf(printheader = True, kkk!uptt, "")
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
rs1.Open "select * from credita where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); " To,  S.L. Code : "; Tab(20); Mid$(rs1!SUBLEDGER, 1, 5); Tab(46); "C/Note No. : "; Trim(rs1!INVOICENO); Tab(74); "Dated     : "; Chr(27) + Chr(72); rs1!InvoiceDate
    If kkk.State = 1 Then kkk.Close
    kkk.Open "select * from sledger where  " & stringyear & " and subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
       Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE; Tab(45); Chr(27) + Chr(71); "Bilty No.  : "; Chr(27) + Chr(72); Trim(rs1!biltyno); Tab(75); Chr(27) + Chr(71); "Dated     : "; Chr(27) + Chr(72); IIf(IsNull(rs1!OrderDate), "  /  /    ", rs1!OrderDate)
       Print #1, Tab(3); IIf(IsNull(kkk!address1), " ", kkk!address1); Tab(45); Chr(27) + Chr(71); "Freight    : "; Chr(27) + Chr(72); Trim(rs1!freight); Tab(75); Chr(27) + Chr(71); "Bundle(s) : "; Chr(27) + Chr(72); Trim(rs1!bundles)
       Print #1, Tab(3); IIf(IsNull(kkk!ADDRESS2), " ", kkk!ADDRESS2); Tab(45); Chr(27) + Chr(71); "Demurrage  : "; Chr(27) + Chr(72); Trim(rs1!marka); Tab(75); Chr(27) + Chr(71); "Weight    : "; Chr(27) + Chr(72); Trim(rs1!weight) & " Kgs."
       Print #1, Tab(3); IIf(IsNull(kkk!ADDRESS3), " ", kkk!ADDRESS3); Chr(27) + Chr(72)
       kkk.Close
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
    If kk.State = 1 Then kk.Close
    kk.Open "select * from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,sno ", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                tdata.Open "Select bookname from books where " & stringyear & "' and bookcode='" + Trim(kk!Bookcode) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
                Print #1, Tab(0); rsets(Trim(str(sno)), 4); Tab(7); Trim(tdata!Bookname); Tab(52); rsets(Trim(str(kk!quantity)), 5); Tab(58); rsets(Trim(Format(str(kk!rate), "0.00")), 8); Tab(68); rsets(Trim(Format(str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!quantity
                Line = Line + 1
                If Line > MaxLine - 5 Then
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
       Print #1, Line; repli("-", 96)
       Line = Line + 1
       Print #1, Tab(52); rsets(Trim(str(totalquantity)), 5); Tab(84); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12)
       Line = Line + 1
       If kk.State = 1 Then kk.Close
       kk.Open "Select * from CreditC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
       If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount - kk!amount
                    Else
                        netamount = netamount + kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(60); Trim(kk!Text) + "    " + Trim(Format(str(myround(kk!rate, 2)), "0.00")); Tab(84); rsets(Trim(Format(str(myround(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(60); Trim(kk!Text); Tab(84); rsets(Trim(Format(str(myround(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
        End If
        Print #1, Tab(84); repli("-", 12)
        Print #1, Chr(27) + Chr(71); Tab(45); "NET AMOUNT Cr. TO YOUR A/C :"; Tab(85); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
        Line = Line + 2
        VNetamt = netamount
        If kk.State = 1 Then kk.Close
        kk.Open "Select * from CreditA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(60); kk!txt1; Tab(84); rsets(Trim(Format(str(Abs(myround(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + myround(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(60); kk!txt2; Tab(84); rsets(Trim(Format(str(Abs(myround(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + myround(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(60); "BY BANK "; Tab(84); rsets(Trim(Format(str(Abs(myround(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - myround(kk!baa, 2)
             End If
        End If
        Print #1, Tab(84); repli("-", 12)
        Line = Line + 1
      ' PRINT THE FOOTER IN INVOICE START
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
        Close #1
        PrintOption.Show
End Sub
  
Sub CREDITCalc()
    'OTHERcredit.calc
     mga.Caption = Format(myround(totalamount, 2), "0.00")
     mgd.Caption = Format(myround(totaldiscount, 2), "0.00")
     mna.Caption = Format(myround((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
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
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
If kk.State = 1 Then
   kk.Close
End If
If Edit = False Then
    kk.Open "SELECT MAX(INVOICENO) FROM creditA where " & stringyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
    If kk(0) > 0 Then
        kk.MoveLast
        Me.I_NO.Text = Trim(str(kk(0) + 1))
    Else
        Me.I_NO.Text = "1"
    End If
    kk.Close
    End If
        Dim ctl As Control
        For Each ctl In Me.Controls
            If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
                If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) Then
                    ctl.Text = ""
                End If
                ctl.Enabled = False
                CRITNOTE.customercode.Enabled = False
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
        Unload OTHERCredit
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
                If RS.State = 1 Then
                    RS.Close
                End If
                RS.Open "select * from books where   " & stringyear & "' ", CON, adOpenKeyset, adLockReadOnly, adCmdText
                Row = Grid1.Row
                Col = Grid1.Col
                If Trim(Grid1.Text) <> "" Then
                    If Not RS.BOF Then
                        RS.MoveFirst
                        RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF Then
                            tempmeb.Visible = True
                            tempmeb.SetFocus
                            RS.Close
                            templost = False
                            Exit Function
                        Else
                            Grid1.Text = RS(0)
                            Grid1.Col = 2
                            Grid1.Text = RS(1)
                         'If Not edit Then
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
                                Set kk = CON.Execute("select DISCATEGORY from sledger where " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
                                Grid1.Col = 6
                                If Grid1.Text = "" Or addmode = True Then
                                
                                 If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    kk.Close
                                    Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
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
                                Grid1.Text = Format(myround(q * r, 2), "0.00")
                                Grid1.Col = 8
                                Grid1.Text = Format(myround((q * r) * (D / 100), 2), "0.00")
                            'End If
                            End If
                            Grid1.Col = Col
                            RS.Close
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
                Grid1.Text = Format(myround(q * r, 2), "0.00")
                Grid1.Col = 8
                Grid1.Text = Format(myround((q * r) * (D / 100), 2), "0.00")
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
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
        Next
        CREDITCalc
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
                        
                           Commandother.SetFocus
                        Exit Sub
                    End If
                End If
                Grid1.Row = Row
                Grid1.Col = Col
                Grid1.Text = Bookname.Text
                '/*************************
                If RS.State = 1 Then
                    RS.Close
                End If
                RS.Open "select * from books where  " & stringyear & "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
                Row = Grid1.Row
                Col = Grid1.Col
                If Trim(Grid1.Text) <> "" Then
                    If Not RS.BOF Then
                        RS.MoveFirst
                        RS.Find "bookname='" + Trim(Grid1.Text) + "'"
                        If RS.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            RS.Close
                            Exit Sub
                        Else
                            
                            Grid1.Col = 1
                            Grid1.Text = RS(0)
                            Grid1.Col = 2
                            Grid1.Text = RS(1)
                        '    If Not edit Then
                                 Grid1.Col = 3
                                If Trim(Grid1.Text) = "" Then
                                        Grid1.Text = 0
                                End If
                                q = Val(Grid1.Text)
                                Grid1.Col = 5
                                Grid1.Text = Format(RS(3), "0.00")
                                r = RS(3)
                                '/******************
                                Set kk = CON.Execute("select DISCATEGORY from sledger where  " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
                                Grid1.Col = 6
                                If Trim(kk(0)) <> "" Then
                                    tempstr = Trim(kk(0))
                                    kk.Close
                                    Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
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
                                Grid1.Text = myround(q * r, 2)
                                Grid1.Col = 8
                                Grid1.Text = myround((q * r) * (D / 100), 2)
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
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
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


Private Sub cmbAgentName_LostFocus()
If cmbAgentName.Text = "" Then
   MsgBox "Enter a Agent Name.. "
   cmbAgentName.SetFocus
   Exit Sub
Else
  Dim rs1 As New ADODB.Recordset
  rs1.Open "select * from agentmaster where " & stringyear & " and agentname='" & cmbAgentName.Text & "' order by agentname", CON, adOpenKeyset, adLockReadOnly, adCmdText
  If rs1.RecordCount <= 0 Then
     MsgBox "Enter valid Agent Name.. "
     cmbAgentName.SetFocus
  End If
End If
End Sub

Private Sub Command1_Click()
printheader = False
printinvoice
End Sub

Private Sub Commandabandon_Click()
CREDITAbandon
SetButton Commandadd, Commandedit, Commandsave, Commanddelete

End Sub
Private Sub Commandadd_Click()
    addmode = True
    Edit = False
    CREDITAbandon
    Dim RS As ADODB.Recordset
    addoredit = True
    addmode = True
    Set RS = New ADODB.Recordset
    Dim TEMPNUM As Integer
    If Edit = False Then
       If CON.Execute("SELECT MAX(INVOICENO) FROM creditA where " & stringyear & "")(0) >= Val(Trim(Me.I_NO.Text)) Then
            Me.I_NO.Text = CON.Execute("SELECT MAX(INVOICENO) FROM creditA where " & stringyear & "")(0) + 1
              RS.Open "select * from tempcritnote where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
              If RS.BOF Then
                 RS.addNew
              End If
              RS!In = Val(Me.I_NO.Text)
              RS.Update
              RS.Close
        End If
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
    Commandsave.Enabled = True
    Commandsearch.Enabled = False
    Grid1.Enabled = True
    Me.customercode.Enabled = True
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
        RS.Close
    End If
    RS.Open "select * from books where   " & stringyear & "  order by bookcode", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
            Set kk = CON.Execute("select DISCATEGORY from sledger where  " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
            Grid1.Col = 6
            If Trim(kk(0)) <> "" Then
                tempstr = Trim(kk(0))
                kk.Close
                Set kk = CON.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
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
            Grid1.Text = Format(myround(q * r, 2), "0.00")
            Grid1.Col = 8
            Grid1.Text = Format(myround((q * r) * (D / 100), 2), "0.00")
            If Not RS.EOF Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Row = Grid1.Row + 1
                RS.MoveNext
            End If
        Loop
        '/**fghfghgh
        '    Grid1.col = col
    End If
    RS.Close
   ' row = Grid1.row
   ' col = Grid1.col
    totalamount = 0
    totaldiscount = 0
    Me.tqu.Caption = ""
    For I = 1 To Grid1.Rows - 1
            Grid1.Row = I
            Grid1.Col = 7
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
            Grid1.Col = 3
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
If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbNo Then
        Exit Sub
Else
                CON.Execute ("DELETE from CREDITA where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("DELETE from CREDITB where INVOICENO = " + Trim(I_NO.Text))
                CON.Execute ("DELETE from CREDITC where INVOICENO = " + Trim(I_NO.Text))
                CREDITAbandon
End If
End Sub

Private Sub Commandedit_Click()
    Commandadd.Enabled = False
    'Me.Commandedit.Enabled = False
    Picture5.Enabled = True
    Commandother.Enabled = True
    Commandadd.Enabled = False
    Commandedit.Enabled = False
    Commandall.Enabled = True
    Commandsave.Enabled = True
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
    CON.Execute ("DELETE from CREDITCtmp")
    CON.Execute ("insert into CREDITCtmp  select * from CREDITC where INVOICENO = " + Trim(I_NO.Text))
    ' CREDITTMP creation end
    addoredit = False
    HIT
    
    OTHERCredit.TOP = 0
    OTHERCredit.Left = 0
    OTHERCredit.Visible = False
   
   
    
    
End Sub
Private Sub Commandother_Click()
    Me.Enabled = False
    OTHERCredit.TOP = 0
    OTHERCredit.Left = 0
    'Unload OTHERcredit
    'Load OTHERcredit
    OTHERCredit.Show
   
End Sub
Private Sub CommandPrint_Click()
printheader = True
printinvoice

End Sub
Private Sub Commandreturn_Click()
    Unload Me
    addoredit = False
End Sub
Private Sub Commandsave_Click()
    Dim SAVED As Boolean
    Dim LAMOUNT As Double
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    If MsgBox("Do you want to save it now ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    SAVED = False
    Grid1.Row = 1
    Grid1.Col = 1
    If Trim(Grid1.Text) = "" Then
        MsgBox "Please Enter item.... "
        Exit Sub
    End If
    If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(customercode.Text) <> "" Then
        If Edit Then
            CON.Execute ("DELETE from CREDITA where INVOICENO = " + Trim(I_NO.Text))
            CON.Execute ("DELETE from CREDITB where INVOICENO = " + Trim(I_NO.Text))
            CON.Execute ("DELETE from CREDITC where INVOICENO = " + Trim(I_NO.Text))
        End If
        If RS.State = 1 Then
            RS.Close
        End If
            LAMOUNT = 0
            RS.Open "select * from credita where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
            If Not Edit Then
again:
               If CON.Execute("SELECT MAX(INVOICENO) FROM creditA where " & stringyear & "")(0) >= Val(Trim(Me.I_NO.Text)) Then
                    'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'GoTo again
               End If
            End If
            RS.addNew
            RS!INVOICENO = Val(Me.I_NO.Text)
            RS!InvoiceDate = Me.i_dt.Text
            RS!Genledger = Trim(Me.Genledger.Text)
            RS!SUBLEDGER = Trim(Me.customercode.Text)
            RS!marka = Trim(Me.marka.Text)
            RS!bundles = Trim(Me.bundles)
            RS!biltyno = Trim(Me.biltno.Text)
            If Trim(Me.bdated) = Trim("__/__/____") Then
                RS!BILTYDATE = Null '"__/__/____"
            Else
                RS!BILTYDATE = Me.bdated
            End If

            RS!freight = Me.freight & ""
            RS!weight = Me.weight & ""
            RS!netamount = myround(Val(Trim(Me.mna.Caption)), 2)
            RS!GAmount = (Me.totalamount - Me.totaldiscount)
            RS!txt1 = Trim(OTHERCredit.T1TEXT.Text)
            RS!txt1a = Val(Trim(OTHERCredit.T1.Text))
            RS!txt2 = Trim(OTHERCredit.T2TEXT.Text)
            RS!txt2a = Val(Trim(OTHERCredit.T2.Text))
            RS!baa = Val(Trim(OTHERCredit.T3TEXT.Text))
            RS!agentname = cmbAgentName.Text
            Dim trs As New ADODB.Recordset
            trs.Open " SELECT DISTCODE  FROM SLEDGER  WHERE  " & stringyear & " and SUBLEDGER='" & customercode.Text & "'", CON, adOpenStatic, adLockOptimistic, adCmdText
            If Not trs.BOF Then
               RS!District = trs!distcode & ""
            Else
               RS!District = ""
            End If
            trs.Close
err1:
            If Not Edit Then
                If CON.Execute("SELECT MAX(INVOICENO) FROM creditA where " & stringyear & "")(0) >= Val(Trim(Me.I_NO.Text)) Then
                    'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
                    'rs!INVOICENO = Val(Me.I_NO.Text)
                    On Error GoTo err1
                End If
            End If
            RS.Update
            On Error GoTo 0
            RS.Close
            RS.Open "select * from creditb where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
            Dim I As Integer
            RRRR = Grid1.Row
            CCCC = Grid1.Col
            For I = 1 To maxrow
                Grid1.Row = I
                Grid1.Col = 1
                If Trim(Grid1.Text) <> "" Then
                    Grid1.Col = 3
                    If Val(Trim(Grid1.Text)) > 0 Then
                        RS.addNew
                        Grid1.Col = 1
                        RS!INVOICENO = Val(Me.I_NO.Text)
                        RS!InvoiceDate = Me.i_dt.Text
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
                        RS!discount = Trim(Grid1.Text)
                        Grid1.Col = 8
                        RS!netamount = LAMOUNT - Trim(Grid1.Text)
                        LAMOUNT = 0
                        RS.Update
                    End If
                End If
            Next
            RS.Close
            Grid1.TopRow = 1
            RS.Open "select * from creditc where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
            '/******
            Dim temprs As ADODB.Recordset
            Set temprs = New ADODB.Recordset
            For I = 1 To OTHERCredit.mrow
                    OTHERCredit.Grid1.Row = I
                    OTHERCredit.Grid1.Col = 0
                    If Trim(OTHERCredit.Grid1.Text) <> "" Then
                        RS.addNew
                        RS!INVOICENO = Val(Me.I_NO.Text)
                        RS!InvoiceDate = Me.i_dt.Text
                        RS!GAmount = (Me.totalamount - Me.totaldiscount)
                        RS!Text = Trim(OTHERCredit.Grid1.Text)
                        If temprs.State = 1 Then
                            temprs.Close
                        End If
                        If Edit Then
                        temprs.Open "select * from creditctmp where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
                        If OTHERCredit.Grid1.Text <> "" Then
                                temprs.Find "TEXT='" + Trim(OTHERCredit.Grid1.Text) + "'"
                                RS!Genledger = temprs!Genledger & ""
                                RS!SUBLEDGER = temprs!SUBLEDGER & ""
                                RS!DebitorCredit = temprs!DebitorCredit & ""
                                RS!RYN = temprs!RYN & ""
                        End If
                        temprs.Close
                        Else
                        temprs.Open "select * from creditend where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
                        If OTHERCredit.Grid1.Text <> "" Then
                           temprs.Find "TEXT='" + Trim(OTHERCredit.Grid1.Text) + "'"
                           RS!Genledger = temprs!Genledger & ""
                           RS!SUBLEDGER = temprs!SUBLEDGER & ""
                           RS!DebitorCredit = temprs!DebitorCredit & ""
                           RS!RYN = temprs!RYN & ""
                        End If
                        temprs.Close
                        End If
                        OTHERCredit.Grid1.Col = 1
                        RS!rate = Val(Trim(OTHERCredit.Grid1.Text))
                        If Val(Trim(OTHERCredit.Grid1.Text)) > 0 Then
                            RS!amount = (Me.totalamount - Me.totaldiscount) * (Val(Trim(OTHERCredit.Grid1.Text)) / 100)
                        Else
                        OTHERCredit.Grid1.Col = 2
                            RS!amount = Val(Trim(OTHERCredit.Grid1.Text))
                        End If
                    RS.Update
                    End If
            Next
            RS.Close
            If addmode = True Then
                RS.Open "select * from tempcritnote where " & stringyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
                If RS.BOF Then
                   RS.addNew
                End If
                RS!In = Val(Me.I_NO.Text)
                RS.Update
                RS.Close
                End If
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
   End Sub

Private Sub Commandsearch_Click()
    Me.Enabled = False
    Call searchscreen.tempr(13, "CREDITITEMNOTE")
End Sub

Private Sub customercode_KeyPress(KeyAscii As Integer)
   ' If KeyAscii = 13 Then
   ' SendKeys "{DOWN}"
   ' SendKeys "{TAB}"
   ' marka.SetFocus
   ' End If
End Sub
Private Sub customercode_LostFocus()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "select * from sledger where  " & stringyear & "  and gledger='" + Genledger.Text + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    RS.Find "subledger='" + Trim(customercode.Text) + "'"
    If RS.EOF Then
        customercode.SetFocus
        HIT
        RS.Close
        Exit Sub
    End If
    
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    If RS!distcode <> "" And addmode = True Then
       rs1.Open "Select * from Districts where  " & stringyear & " and Districtname = '" & RS!distcode & "'", CON, adOpenStatic, adLockReadOnly
       If rs1.RecordCount > 0 Then
          Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
       End If
    End If
    If RS.State = 1 Then RS.Close
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


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If Grid1.Row >= 1 Then
           Grid1.RemoveItem Grid1.Row - 1
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
           If addmode = True Then
                SendKeys "{DOWN}"

           End If
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
SetButton Commandadd, Commandedit, Commandsave, Commanddelete



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

    Grid1.TOP = 1470
   ' Set CON = New ADODB.Connection
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Me.TOP = 0
    Me.Left = 0
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
    Grid1.ColWidth(0) = 200
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 750
    Grid1.ColWidth(4) = 750
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 1200
    Grid1.ColWidth(8) = 1200
    Me.CommandPrint.Enabled = True
    
    
    RS.Open "select * from books where  " & stringyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            Me.Bookcode.AddItem RS(0)
            Me.Bookname.AddItem IIf(IsNull(RS(1)), "", RS(1))
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    
    RS.Close
    Genledger.AddItem "SUNDRY DEBTORS"
    Genledger.Text = "SUNDRY DEBTORS"
    RS.Open "select SUBLEDGER from sledger where  " & stringyear & " and gledger='" + Trim(Genledger.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            Me.customercode.AddItem RS(0)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    

     '*******Agent  combo fill
    RS.Open "select Distinct Agentname from AgentMaster where " & stringyear & " order by agentname", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
    RS.Open "select * from tempcritnote where " & stringyear, CON, adOpenStatic, adLockOptimistic, adCmdText
    If Not RS.BOF Then
       Me.I_NO.Text = RS!In
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
       kk.Open "SELECT MAX(INVOICENO) FROM creditA where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
       If kk(0) <> "" Then
            Me.I_NO.Text = Trim(str(kk(0) + 1))
       Else
            Me.I_NO.Text = "1"
       End If
        kk.Close
   End If
   RS.Close
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
                tempmeb.Left = Grid1.CellLeft + 80
                tempmeb.TOP = Grid1.TOP + Grid1.CellTop '- 50
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.Text = Grid1.Text
                Bookname.TOP = Grid1.TOP + Grid1.CellTop
                Bookname.Left = Grid1.CellLeft + 80
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
       i_dt.Enabled = True
        i_dt.SetFocus
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
Dim RS As ADODB.Recordset
Dim rs3  As ADODB.Recordset
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
Set RS = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
If Trim(I_NO.Text) = "" Then
        MsgBox "Credit Note no cannot be null"
        I_NO.SetFocus
Else
    If RS.State = 1 Then RS.Close
    
    'rs.Open "select * from credita where " & stringyear, con, adOpenKeyset, adLockReadOnly, adcmdtext
    RS.Open "Select * from  CREDITA where INVOICENO = " + Trim(I_NO.Text) + "", CON, adOpenStatic, adLockReadOnly
    RS.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
    rs3.Open "Select * from cnf1a  where cnn = " & Val(I_NO.Text) & "", CON, adOpenStatic, adLockReadOnly, adCmdText
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
   Dim ctl As Control
   For Each ctl In Me.Controls
       If Not TypeOf ctl Is CommandButton Then
          ctl.Enabled = True
       End If
   Next
        Me.Commandother.Enabled = True
        I_NO.Text = RS!INVOICENO
        Me.i_dt.Text = RS!InvoiceDate
        Me.Genledger.Text = Trim(RS!Genledger)
        Me.customercode.Text = Trim(RS!SUBLEDGER)
        Me.textbox.Text = Trim(RS!SUBLEDGER)
        Me.marka.Text = Trim(RS!marka)
        Me.bundles = Trim(RS!bundles)
        Me.biltno.Text = Trim(RS!biltyno)
        
        If IsNull(RS!BILTYDATE) Then
           Me.bdated = "__/__/____"
        Else
           Me.bdated = RS!BILTYDATE
        End If
        Me.freight = Trim(RS!freight)
        Me.weight = Trim(RS!weight)
        Me.labelbybank = Trim(RS!baa)
        mna.Caption = RS!netamount
        Me.cmbAgentName.Text = IIf(IsNull(RS!agentname), "", RS!agentname)
        RS.Close
'*/**/*/*/*/*//*/*
        If RS.State = 1 Then
            RS.Close
        End If
 RS.Open "Select * from CREDITB where INVOICENO =" + Trim(I_NO.Text) + "", CON, adOpenStatic, adLockReadOnly
       '' rs.Find "INVOICENO='" + Trim(I_NO.Text) + "'"
        Grid1.TopRow = 1
        If Not RS.EOF Then
            Grid1.Row = 1
            Grid1.Col = 1
            Do While Not RS.EOF
               If Trim(RS!INVOICENO) = Trim(I_NO.Text) Then
               Grid1.Col = 1
                Grid1.Text = Trim(RS!Bookcode)
                If kk.State = 1 Then
                    kk.Close
                End If
                kk.Open "select * from books where   " & stringyear & " and bookcode='" + Trim(RS!Bookcode) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
                Grid1.Col = 2
                Grid1.Text = Trim(kk!Bookname)
                Grid1.Col = 3
                Grid1.Text = Trim(RS!quantity)
                Grid1.Col = 5
                Grid1.Text = Format(myround(RS!rate, 2), "0.00")
                Grid1.Col = 7
                Grid1.Text = Format(myround(RS!amount, 2), "0.00")
                Grid1.Col = 4
                Grid1.Text = Format(myround(RS!PRINTORDER, 2), "0.00")
                Grid1.Col = 6
                Grid1.Text = Format(myround(RS!discount, 2), "0.00")
                Grid1.Col = 8
                Grid1.Text = Format(myround(RS!amount * (RS!discount / 100), 2), "0.00")
                End If
                If Not RS.EOF Then
                    RS.MoveNext
                End If
                Grid1.Row = Grid1.Row + 1
                Grid1.Rows = Grid1.Rows + 1
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
            totalamount = totalamount + Val(Trim(Grid1.Text))
            Grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
        Next
     mga.Caption = Format(myround(totalamount, 2), "0.00")
     mgd.Caption = Format(myround(totaldiscount, 2), "0.00")
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
                    If RS.State = 1 Then
                        RS.Close
                    End If
                    RS.Open "select * from books where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
                    If Not RS.BOF Then
                        RS.MoveFirst
                        RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF And Trim(Grid1.Text) <> "" Then
                            RS.Close
                            Exit Sub
                        Else
                            RS.Close
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
              '       SendKeys "{LEFT}"
                    Grid1.SetFocus
                    Grid1_Click
                Case 5
                    If Val(tempmeb.Text) > 0 Then
                        Grid1.Col = Grid1.Col - 1
                        Grid1.SetFocus
                        Grid1_Click
                    End If
                Case 6
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
    
    Open "" + App.Path + "\vipin.txt" For Output As #1
    Line = 0
header:
    If kkk.State = 1 Then
        kkk.Close
    End If
    CNSetup
    kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If flagyes = True Then

    
      If Not kkk.BOF Then
        Print #1, Chr(27) + Chr(15) + Chr(14)
        Print #1, Tab(T1); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname)) '
        Print #1, Tab(T2 - 7); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
        Print #1, Tab(T3); Trim(kkk!phone1)
        Line = Line + 4
    End If
  
    Print #1, repli("-", 150)

  End If
    If rs1.State = 1 Then
        rs1.Close
    End If
    rs1.Open "select * from credita where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
        Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!SUBLEDGER; Tab(T5); "Credit Note No. : "; Trim(rs1!INVOICENO); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!InvoiceDate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.Close
            End If
            kkk.Open "select * from sledger where  " & stringyear & " and subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
            If Not kkk.EOF Then
                Print #1, Tab(3); kkk!DESCFORINVOICE
                'Print #1, Tab(3); kkk!ADDRESS1; Tab(t5); "Order by : "; Trim(rs1!ORDERBY); Tab(t8 + 5); "Dt. "; Tab(t8 + 12); rs1!ORDERDATE
                Print #1, Tab(3); kkk!distcode; Tab(T5); "Bilty No.: "; Trim(rs1!biltyno); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!BILTYDATE
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
            kk.Open "select * from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                        tdata.Open "select sum(amount) from CREDITB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
           kk.Open "Select * from CREDITC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
           If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
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
           kk.Open "Select * from CREDITA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
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
Open "" + App.Path + "\vipin.txt" For Output As #1
Line = 0
header:
If kkk.State = 1 Then
      kkk.Close
End If
 CNSetup
 kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
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


Print #1, Chr(27) + Chr(15) + Chr(14); Tab(25); dspace(Trim("CRDEIT NOTE ")); Chr(20); Tab(T4 + 6); IIf(printheader = True, kkk!uptt, "")
If printheader = True Then
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
rs1.Open "select * from credita where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To, S.L. Code : "; Tab(T1 + 8); Mid$(rs1!SUBLEDGER, 1, 5); Tab(T5); "C/Note No. : "; Trim(rs1!INVOICENO); Tab(T8); "Dated     : "; rs1!InvoiceDate
    kine = libe + 1
    If kkk.State = 1 Then
        kkk.Close
    End If
    kkk.Open "select * from sledger where  " & stringyear & " and subledger='" + Trim(rs1!SUBLEDGER) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(3); "M/s " & kkk!DESCFORINVOICE; Tab(T5); "Bilty No.  : "; Trim(rs1!biltyno); Tab(T8); "Dated     : "; IIf(IsNull(rs1!OrderDate), "  /  /    ", rs1!OrderDate)
        Print #1, Tab(3); kkk!address1; Tab(T5); "Freight    : "; Trim(rs1!freight); Tab(T8); "Bundle(s) :  "; Trim(rs1!bundles);
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
    kk.Open "select * from CreditB where invoiceno=" + Trim(rs1!INVOICENO) + " order by discount,printorder", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
                tdata.Open "select sum(amount) from CreditB where invoiceno=" + Trim(rs1!INVOICENO) + " and discount=" + Trim(str(cdiscount)) + " group by discount", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
        Print #1, repli("-", 150)
        Print #1, Tab(T5 - 4); rsets(Trim(str(totalquantity)), 5); Tab(T8 + 5); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12)
       'by vk Print #1, Tab(t8); repli("-", 22)
        Line = Line + 2
        If kk.State = 1 Then
             kk.Close
        End If
        kk.Open "Select * from CreditC where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            Do While Not kk.EOF
                If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(T5 + 10); Trim(kk!Text) + "    " + Trim(Format(str(myround(kk!rate, 2)), "0.00")); Tab(T8 + 5); rsets(Trim(Format(str(myround(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5 + 10); Trim(kk!Text); Tab(T8 + 5); rsets(Trim(Format(str(myround(kk!amount, 2)), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
            Print #1, Tab(T8); repli("-", 22)
            Print #1, Chr(27) + Chr(71); Tab(T5 - 10); "NET AMOUNT Cr. TO YOUR A/C :"; Tab(T8 + 6); rsets(Trim(Format(str(myround(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
            Line = Line + 2
            VNetamt = netamount
        End If
        kk.Close
        kk.Open "Select * from CreditA where invoiceno=" + Trim(Me.I_NO.Text), CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(T5 + 10); kk!txt1; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!txt1a, 2))), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + myround(kk!txt1a, 2)
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(T5 + 10); kk!txt2; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!txt2a, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + myround(kk!txt2a, 2)
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(T5 + 10); "BY BANK "; Tab(T8 + 5); rsets(Trim(Format(str(Abs(myround(kk!baa, 2))), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - myround(kk!baa, 2)
             End If
        End If
        Print #1, Tab(T8); repli("-", 22)
        'Print #1, Tab(T5 + 10); Chr(27) + Chr(71); "BALANCE  : "; Tab(T8 + 6); rsets(Trim(Format(Str(myround(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72)
       ' Print #1, Tab(T8); repli("-", 22)
       Line = Line + 1
       ' PRINT THE FOOTER IN INVOICE START
        Do While Line < 61
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

