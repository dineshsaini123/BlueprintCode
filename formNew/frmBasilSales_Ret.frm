VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBasilSales_Ret 
   ClientHeight    =   7956
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10788
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7956
   ScaleWidth      =   10788
   Begin VB.Frame panel 
      Caption         =   "Basil Sales (Return)"
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
      Height          =   7890
      Left            =   15
      TabIndex        =   8
      Top             =   60
      Width           =   10680
      Begin VB.ComboBox cboCatII1 
         Height          =   315
         Left            =   1785
         TabIndex        =   6
         Top             =   1320
         Width           =   795
      End
      Begin VB.ComboBox cboCatII 
         Height          =   315
         Left            =   975
         TabIndex        =   5
         Top             =   1320
         Width           =   795
      End
      Begin VB.ComboBox txtMark 
         Height          =   315
         ItemData        =   "frmBasilSales_Ret.frx":0000
         Left            =   3960
         List            =   "frmBasilSales_Ret.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   765
      End
      Begin VB.ComboBox cmbareaname 
         Appearance      =   0  'Flat
         Height          =   1296
         Left            =   7620
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.ComboBox cmbAgentName 
         Height          =   315
         Left            =   3360
         TabIndex        =   22
         Top             =   6405
         Width           =   2520
      End
      Begin VB.ComboBox cmbdiscountcat 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   870
      End
      Begin VB.ComboBox Combosldistrictcode 
         Height          =   315
         Left            =   2580
         TabIndex        =   7
         Top             =   1320
         Width           =   1365
      End
      Begin VB.Frame frame1 
         Height          =   675
         Left            =   120
         TabIndex        =   21
         Top             =   270
         Width           =   2010
         Begin VB.OptionButton Optioncredit 
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
            Left            =   1020
            TabIndex        =   1
            Top             =   270
            Width           =   885
         End
         Begin VB.OptionButton Optioncash 
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
            Left            =   90
            TabIndex        =   0
            Top             =   270
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.CommandButton Commandother 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&End Part"
         Height          =   510
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6165
         Width           =   1125
      End
      Begin VB.ComboBox Bookname 
         Height          =   912
         Left            =   3600
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   19
         Top             =   2880
         Width           =   2295
      End
      Begin VB.ComboBox Bookcode 
         Height          =   720
         Left            =   540
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   18
         Top             =   2880
         Width           =   2355
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   225
         ScaleHeight     =   792
         ScaleWidth      =   9888
         TabIndex        =   11
         Top             =   6930
         Width           =   9885
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edi&t"
            Height          =   690
            Left            =   1125
            Picture         =   "frmBasilSales_Ret.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Height          =   690
            Left            =   2220
            Picture         =   "frmBasilSales_Ret.frx":0446
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   690
            Left            =   3330
            Picture         =   "frmBasilSales_Ret.frx":102A
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Enabled         =   0   'False
            Height          =   690
            Left            =   4425
            Picture         =   "frmBasilSales_Ret.frx":15B4
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   690
            Left            =   5520
            Picture         =   "frmBasilSales_Ret.frx":2198
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   690
            Left            =   8775
            Picture         =   "frmBasilSales_Ret.frx":2D7C
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton Commandadd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   690
            Left            =   30
            Picture         =   "frmBasilSales_Ret.frx":3960
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton Commandprintnh 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Enabled         =   0   'False
            Height          =   690
            Left            =   6615
            Picture         =   "frmBasilSales_Ret.frx":4544
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton cmdSalep 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sale Estimate"
            Enabled         =   0   'False
            Height          =   690
            Left            =   7710
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton CommandPrint 
            Caption         =   "&Print"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2475
            TabIndex        =   16
            Top             =   810
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CommandButton Commandhelp 
            Caption         =   "Help"
            Height          =   375
            Left            =   2565
            TabIndex        =   12
            Top             =   810
            Visible         =   0   'False
            Width           =   915
         End
      End
      Begin VB.ComboBox customercode 
         Appearance      =   0  'Flat
         Height          =   1296
         Left            =   6885
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.ComboBox Genledger 
         Height          =   315
         Left            =   10305
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   330
      End
      Begin MSMask.MaskEdBox textbox 
         Height          =   315
         Left            =   6900
         TabIndex        =   3
         Top             =   375
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   572
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   4140
         Left            =   120
         TabIndex        =   17
         Top             =   1890
         Width           =   10110
         _ExtentX        =   17844
         _ExtentY        =   7303
         _Version        =   393216
         BackColorFixed  =   7917545
         BackColorBkg    =   16777215
         FillStyle       =   1
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox bundles 
         Height          =   345
         Left            =   5370
         TabIndex        =   15
         Top             =   1290
         Width           =   1515
         _ExtentX        =   2667
         _ExtentY        =   593
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_OB 
         Height          =   345
         Left            =   4740
         TabIndex        =   14
         Top             =   1290
         Width           =   585
         _ExtentX        =   1037
         _ExtentY        =   593
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
         Height          =   300
         Left            =   3450
         TabIndex        =   2
         Top             =   660
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tempmeb 
         Height          =   285
         Left            =   600
         TabIndex        =   24
         Top             =   2370
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1566
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rate 
         Height          =   285
         Left            =   630
         TabIndex        =   25
         Top             =   3510
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3239
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox amount 
         Height          =   285
         Left            =   720
         TabIndex        =   26
         Top             =   2970
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3260
         _ExtentY        =   487
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox I_NO 
         Height          =   285
         Left            =   3450
         TabIndex        =   27
         Top             =   360
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   508
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "O.No"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4740
         TabIndex        =   47
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label LBLDIS2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(III)"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1785
         TabIndex        =   46
         Top             =   1050
         Width           =   750
      End
      Begin VB.Label lbldis1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(II)"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   930
         TabIndex        =   45
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Godown"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4020
         TabIndex        =   44
         Top             =   1050
         Width           =   660
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Agent :"
         Height          =   195
         Left            =   2010
         TabIndex        =   43
         Top             =   6420
         Width           =   555
      End
      Begin VB.Label lbldis 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount (I)"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   1050
         Width           =   870
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "District Name"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2625
         TabIndex        =   41
         Top             =   1050
         Width           =   1320
      End
      Begin VB.Label labelbybanklbl 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " By Cash : "
         Height          =   255
         Left            =   450
         TabIndex        =   40
         Top             =   6390
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label mgd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         Height          =   255
         Left            =   8175
         TabIndex        =   39
         Top             =   6120
         Width           =   1110
      End
      Begin VB.Label mna 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         Height          =   255
         Left            =   8175
         TabIndex        =   38
         Top             =   6420
         Width           =   1110
      End
      Begin VB.Label mga 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         Height          =   255
         Left            =   6900
         TabIndex        =   37
         Top             =   6120
         Width           =   1200
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dated : "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2265
         TabIndex        =   36
         Top             =   660
         Width           =   570
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Estimate No. : "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2265
         TabIndex        =   35
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Code : "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5880
         TabIndex        =   34
         Top             =   420
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "  Net Amount : "
         Height          =   255
         Left            =   6990
         TabIndex        =   33
         Top             =   6420
         Width           =   1200
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Gross Amount : "
         Height          =   255
         Left            =   6780
         TabIndex        =   32
         Top             =   3720
         Width           =   1200
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5325
         TabIndex        =   31
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Total Discount : "
         Height          =   255
         Left            =   8010
         TabIndex        =   30
         Top             =   3210
         Width           =   1290
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   " Total Quantity : "
         Height          =   255
         Left            =   2010
         TabIndex        =   29
         Top             =   6120
         Width           =   1350
      End
      Begin VB.Label tqu 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0078CFE9&
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   28
         Top             =   6120
         Width           =   885
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   915
         Left            =   180
         Top             =   6870
         Width           =   10005
      End
   End
   Begin Crystal.CrystalReport cr 
      Left            =   10800
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmBasilSales_Ret"
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
Dim category As String
Dim addmode As Boolean
Dim Printheader As Boolean
Dim addoredit As Boolean

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
paperWidth = 81
MaxLine = 72 '60
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
    kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
    If FooterYes = True Then
        Print #1, Tab(0); repli("-", 76)
        Line = Line + 1
        If Line > MaxLine - 8 Then
            Do While Line < 72    '61
                Print #1, ""
                Line = Line + 1
            Loop
        End If
        Line = 0
        LEFTM = 5
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
End If
Print #1, Chr(27) + Chr(77)
Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "                     ", ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("ESTIMATE SALES")))) / 2 - 8); Chr(14); "ESTIMATE SALES"; Chr(20)
Line = Line + 2
Print #1, repli("-", 76)
Line = Line + 1
If rs1.State = 1 Then rs1.close
rs1.Open "casha_basilRet", con, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"

If Not rs1.EOF Then
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        Print #1, Tab(5); "M/s " & IIf(Optioncash.value = True, rs1!CASHPARTYNAME, kkk!DESCFORINVOICE); ; Tab(38); Chr(27) + Chr(71); "Estimate No: "; Chr(27) + Chr(72); Trim(rs1!invoiceNo); Tab(62); Chr(27) + Chr(71); "Dt.:"; Chr(27) + Chr(72); IIf(IsNull(rs1!invoiceDate), "  /  /    ", rs1!invoiceDate)
        Print #1, Tab(5); IIf(IsNull(kkk!address1), "", kkk!address1); Tab(38); Chr(27) + Chr(71); "Agent Name : "; Chr(27) + Chr(72); Trim(rs1!agentname)
        Print #1, Tab(5); IIf(IsNull(kkk!address2), "", kkk!address2); Tab(53); Chr(27) + Chr(71); "(" & txtMark & ")"; Chr(27) + Chr(72)
        Print #1, Tab(5); IIf(IsNull(kkk!address3), "", kkk!address3); Tab(38); Chr(27) + Chr(71); "Remarks    : "; Chr(27) + Chr(72); Trim(rs1!bundles)
        kkk.close
        Print #1, Chr(27) + Chr(71); repli("-", 76); Chr(27) + Chr(72)
        Print #1, Tab(0); Tab(6); Chr(27) + Chr(71); "Book Name"; Tab(39); "Qty."; Tab(47); "Rate"; Tab(56); "Amount"; Tab(66); "Net Amount"; Chr(27) + Chr(72)
        Print #1, Chr(27) + Chr(71); repli("-", 76); Chr(27) + Chr(72)
        Line = Line + 7
    End If
    If called1 Then
        GoTo printagain1
    End If
    If called2 Then
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.close
    kk.Open "select * from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " order by printorder,sno ", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                
                Print #1, Tab(0); Tab(1); Trim(tdata!Bookname); Tab(36); rsets(Trim(Str(kk!QUANTITY)), 5); Tab(43); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(51); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                totalquantity = totalquantity + kk!QUANTITY
                If Len(tdata!Bookname) >= 36 Then
                Line = Line + 2
                Else
                Line = Line + 1
                End If
                
                If Line > MaxLine - 8 Then
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
            If Line > MaxLine - 8 Then
                    called2 = True
                    Pno = Pno + 1
                    FooterYes = True
                    GoTo header
                    
                    
printagain2:
                    called2 = False
                End If
                Print #1, Tab(52); repli("-", 12)
                Line = Line + 1
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and printorder =" + Trim(Str(cdiscount)) + " group by printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(51); rsets(Trim(Format(Str(tdata(0)), "0.00")), 12)
                   Print #1, Tab(25); "Less Discount @ " + Trim(Format(Str(vdis), "0.00")) + " %"; Tab(51); rsets(Trim(Format(tdata!sumdis, "0.00")), 12); Tab(64); rsets(Trim(Format(Str(tdata!sumamt - tdata!sumdis), "0.00")), 12)
                   Print #1, Tab(52); repli("-", 12)
                   Line = Line + 3
                   netamount = netamount + tdata!sumamt - tdata!sumdis
                End If
                tdata.close
             Loop
         End If
    End If
    Print #1, repli("-", 76)
    Print #1, Tab(34); rsets(Trim(Str(totalquantity)), 7); Tab(64); rsets(Trim(Format(Str(netamount), "0.00")), 12)
    Line = Line + 2
    If kk.State = 1 Then kk.close
    kk.Open "Select * from cashc_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
       Do While Not kk.EOF
             If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(43); Trim(kk!Text) + "    " + Trim(Format(Str(kk!rate), "0.00")); Tab(64); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(43); Trim(kk!Text); Tab(64); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    End If
                    Line = Line + 1
                End If
                If Not kk.EOF Then
                    kk.MoveNext
                End If
            Loop
          
        End If
        Print #1, Tab(64); repli("-", 12)
        Print #1, Chr(27) + Chr(71); Tab(44); "NET AMOUNT  : "; Tab(65); rsets(Trim(Format(Str(netamount), "0.00")), 12); Chr(27) + Chr(72)
        VNetamt = netamount
        Line = Line + 2
        kk.close
        kk.Open "Select * from casha_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kk.BOF Then
            If kk!txt1a <> 0 Then
                Print #1, Tab(48); kk!txt1 & "    :"; Tab(69); rsets(Trim(Format(Str(Abs(kk!txt1a)), "0.00")), 12)
                Line = Line + 1
                netamount = netamount + kk!txt1a
             End If
             If kk!txt2a <> 0 Then
                 Print #1, Tab(43); kk!txt2 & " :"; Tab(64); rsets(Trim(Format(Str(Abs(kk!txt2a)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount + kk!txt2a
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(48); "CASH RECD.  :"; Tab(69); rsets(Trim(Format(Str(Abs(kk!baa)), "0.00")), 12)
                 Line = Line + 1
                 netamount = netamount - kk!baa
             End If
             If kk!baa <> 0 Then
                 Print #1, Tab(64); repli("-", 12)
                 Print #1, Tab(48); Chr(27) + Chr(71); "BALANCE     : "; Tab(70); rsets(Trim(Format(Str(Round(netamount, 2)), "0.00")), 12); Chr(27) + Chr(72);
                 Line = Line + 2
              End If
        End If
        Print #1, Tab(64); repli("-", 12)
          Line = Line + 1
        'PRINT THE FOOTER IN INVOICE START
        Print #1, repli("-", 76)
        Line = Line + 1
        Do While Line < 72    '61
            Print #1, ""
            Line = Line + 1
        Loop
        '''Print #1, Tab(0); Chr(27) + Chr(71); toword(Round(VNetamt, 2)); Chr(27) + Chr(72)
        
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        CNSetup
        tempdata.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
        'Print #1, ""
        'Print #1, ""
        'Print #1, ""
        'PRINT THE FOOTER IN INVOICE END
        Close #1
        PrintOption.Show
End Sub

Sub printinvoice_esitimate()
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
    kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
    If FooterYes = True Then
        If Line > MaxLine - 10 Then
            Do While Line < 61
                Print #1, ""
                Line = Line + 1
            Loop
        End If
        Line = 0
        LEFTM = 5
        'Print #1, Tab(0); repli("-", 81)
        'Print #1, Tab(1); "E.& O.E"
        'Print #1, Tab(1); kkk!COURT; Tab(LEFTM + (paperWidth - ((Len(kkk!COURT) + Len(kkk!Cname)) * 0.75))); "FOR " + Trim(kkk!Cname)
        'Print #1, Tab(1); "E.& O.E"
        'Print #1, Tab(1); kkk!COURT; Tab(50); "FOR " + Trim(kkk!CNAME)
        'Print #1, ""
        'Print #1, Tab(1); Chr(27) + Chr(71); "Continued on Page : " & Pno; Chr(27) + Chr(72)
        Print #1, Tab(0); repli("-", 81)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
'        Print #1, ""
        
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
'If FooterYes = True Then
'   Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72)
'End If
Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "                     ", ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("ESTIMATE SALES")))) / 2 - 8); Chr(14); "***ESTIMATE SALES***"; Chr(20)
Line = Line + 1
If Printheader = True Then
   Print #1, Tab(48); kkk!uptt
   Print #1, Tab(48); kkk!cst
   Line = Line + 2
End If
If Printheader = False Then
   Print #1, ""
   Print #1, ""
   Line = Line + 2
End If
Print #1, repli("-", 81)
Line = Line + 1

If rs1.State = 1 Then rs1.close
rs1.Open "casha_basilRet", con, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"

If Not rs1.EOF Then
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kkk.EOF Then
        
      
        Print #1, Tab(5); "M/s " & IIf(Optioncash.value = True, rs1!CASHPARTYNAME, kkk!DESCFORINVOICE); ; Tab(40); Chr(27) + Chr(71); "Estimate No: "; Chr(27) + Chr(72); Trim(rs1!invoiceNo); Tab(62); Chr(27) + Chr(71); "Dt.:"; Chr(27) + Chr(72); IIf(IsNull(rs1!invoiceDate), "  /  /    ", rs1!invoiceDate)
        Print #1, Tab(5); IIf(IsNull(kkk!address1), "", kkk!address1); Tab(40); Chr(27) + Chr(71); "Remarks    : "; Chr(27) + Chr(72); Trim(rs1!bundles)
        Print #1, Tab(5); IIf(IsNull(kkk!address2), "", kkk!address2); Tab(40); Chr(27) + Chr(71); "Agent Name : "; Chr(27) + Chr(72); Trim(rs1!agentname)
        Print #1, Tab(5); IIf(IsNull(kkk!address3), "", kkk!address3)
        kkk.close
        Print #1, Chr(27) + Chr(71); repli("-", 81)
        Print #1, Tab(0); Tab(6); "Book Description"; Tab(44); "Qty."; Tab(52); "Rate"; Tab(61); "Amount"; Tab(71); "Net Amount"
        Print #1, repli("-", 81); Chr(27) + Chr(72)
        Line = Line + 7
    End If
    If called1 Then
        GoTo printagain1
    End If
    If called2 Then
        GoTo printagain2
    End If
    If kk.State = 1 Then kk.close
    kk.Open "select * from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " order by printorder,sno ", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                Print #1, Tab(0); Tab(6); Trim(tdata!Bookname); Tab(41); rsets(Trim(Str(kk!QUANTITY)), 5); Tab(48); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(56); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and printorder =" + Trim(Str(cdiscount)) + " group by printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
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
    kk.Open "Select * from cashc_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
       Do While Not kk.EOF
             If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(48); Trim(kk!Text) + "    " + Trim(Format(Str(kk!rate), "0.00")); Tab(69); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(48); Trim(kk!Text); Tab(69); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
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
        kk.Open "Select * from casha_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
        'Print #1, Tab(0); Chr(27) + Chr(71); toword(Round(VNetamt, 2)); Chr(27) + Chr(72)
        'Print #1, Tab(0); repli("-", 81)
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        CNSetup
        tempdata.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
        'Print #1, Tab(1); "E.& O.E"
        'Print #1, Tab(1); tempdata!COURT; Tab(50); "FOR " + Trim(tempdata!CNAME)
        Print #1, Tab(0); repli("-", 81)
        Print #1, ""
        Print #1, ""
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

Sub invoicecalc()
'OTHERCASH.calc
     mga.Caption = Format(Round(totalamount, 2), "0.00")
     mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
     mna.Caption = Format(Round((totalamount - totaldiscount + otheramount - otherdiscount), 2), "0.00")
End Sub
Sub invoiceabandon()
        On Error Resume Next
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
     
        Dim ctl As Control
        For Each ctl In Me.Controls
            If TypeOf ctl Is MaskEdBox Or TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Then
                If UCase(Trim(ctl.Name)) <> UCase(Trim("I_NO")) And UCase(Trim(ctl.Name)) <> UCase(Trim("Genledger")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DTOB")) And UCase(Trim(ctl.Name)) <> UCase(Trim("I_DT")) And UCase(Trim(ctl.Name)) <> UCase(Trim("bdated")) And UCase(Trim(ctl.Name)) <> UCase(Trim("cboRet")) Then
                    ctl.Text = ""
                End If
                ctl.Enabled = False
            End If
        Next
        For I = 1 To maxrow
           grid1.Row = I
            For J = 1 To 8
                grid1.Col = J
               grid1.Text = ""
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
    Dim Row, Col As Integer
    Dim RRR, CCC As Integer
    Dim r, q, D As Double
    Dim mprevcol As Integer
    Dim mq As Currency, mr As Currency, mrot As Currency
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    
      
    On Error GoTo save1:
    
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
                grid1.Text = tempmeb.Text
                '/*************************
                'If RS.State = 1 Then
                '    RS.Close
                'End If
                'RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                Set RS = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(grid1.Text) & "'")
                Row = grid1.Row
                Col = grid1.Col
                If Trim(grid1.Text) <> "" Then
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
                            grid1.Text = RS(0)
                            grid1.Col = 2
                            grid1.Text = RS(1)
                         '   If Not edit Then
                                grid1.Col = 3
                                If Trim(grid1.Text) = "" Then
                                    grid1.Text = 0
                                End If
                                q = Val(grid1.Text)
                                grid1.Col = 5
                                If Trim(grid1.Text) = "" Then
                                grid1.Text = Format(RS(3), "0.00")            'rs(3)
                                r = RS(3)
                                End If
                                '/******************
                                
                                
                       '------------------------------------------
                            category = returnCategory(Trim(RS(2)))
                        If Optioncash.value = True Then
                            
                            If category = "C1" Then
                               Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                            ElseIf category = "C2" Then
                               Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                            ElseIf category = "C3" Then
                               Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII1.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                            End If
                            
                         Else
                            
                            If category = "C1" Then
                               Set kk = con.Execute("select DISCATEGORY from sledger where " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
                            ElseIf category = "C2" Then
                               Set kk = con.Execute("select CATEGORY2 from sledger where " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
                            ElseIf category = "C3" Then
                               Set kk = con.Execute("select CATEGORY3 from sledger where " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
                            End If
                                
                                
                         
                         
                         End If
                         '-----------------------------------
         
                                
                               If kk.BOF Then
                                  GoTo abc
                               End If
                                
                                
                                
                                
                                
                                grid1.Col = 6
                                If grid1.Text = "" And addmode = True Then
                                    If Trim(kk(0)) <> "" Then
                                        tempstr = Trim(kk(0))
                                        
              
                                        kk.close
                                        If category = "C1" Then
                                           If Optioncash.value = True Then
                                               Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                           Else
                                                Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                           End If
                                        
                                        ElseIf category = "C2" Then
                                        
                                           If Optioncash.value = True Then
                                               Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                           Else
                                                Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                           End If
                                        
                                        ElseIf category = "C3" Then
                                        
                                           If Optioncash.value = True Then
                                               Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII1.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                           Else
                                                Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                           End If
                                        
                                        End If

                                        
                                        
                                        '==============================
                                        grid1.Col = 4
                                        If kk.BOF Then
                                             GoTo abc
                                        End If
                                        grid1.Text = Format(kk(0), "0.00")
                                        grid1.Col = 6
                                        grid1.Text = Format(kk(0), "0.00")
                                        D = kk(0)
                                        r = RS(3)
                                     Else
abc:
                                        grid1.Col = 4
                                        grid1.Text = Format(RS(4), "0.00")
                                        grid1.Col = 6
                                        grid1.Text = Format(RS(4), "0.00")
                                        D = RS(4)
                                    End If
                                    
                                    grid1.Col = 7
                                    grid1.Text = Format(Round(q * r, 2), "0.00")
                                    grid1.Col = 8
                                    grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
                            Else
                                                    
                            If grid1.Text = "" And addmode = False Then
                                If (Trim(kk(0)) <> "") Then
                                    tempstr = Trim(kk(0))
                                    kk.close
                                    If Optioncash.value = 0 Then
                                       Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    
                                    Else
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                    grid1.Col = 4
                                    If kk.BOF Then
                                        GoTo abc
                                    End If
                                    grid1.Text = Format(kk(0), "0.00")
                                    grid1.Col = 6
                                    grid1.Text = Format(kk(0), "0.00")
                                    D = kk(0)
                                    r = RS(3)
                            
                                Else
                                    
                                    grid1.Col = 4
                                    grid1.Text = Format(RS!discount, "0.00")
                                    grid1.Col = 6
                                    grid1.Text = Format(RS!discount, "0.00")

                                
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
                    grid1.Text = Format(Trim(tempmeb.Text), "0.00")
                Else
                    grid1.Text = Format(Trim(tempmeb.Text), "0")
                End If
                If Trim(grid1.Text) = "" Then
                    grid1.Text = 0
                End If
                Row = grid1.Row
                Col = grid1.Col
                grid1.Col = 3
                q = Val(Trim(grid1.Text))
                grid1.Col = 5
                r = Val(Trim(grid1.Text))
                grid1.Col = 6
                D = Val(Trim(grid1.Text))
                grid1.Col = 7
                grid1.Text = Format(Round(q * r, 2), "0.00")
                grid1.Col = 8
                grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
                grid1.Col = Col
            Case 4
                grid1.Text = tempmeb.Text
                If Trim(grid1.Text) = "" Then
                    grid1.Text = 0
                End If
        End Select
        Row = grid1.Row
        Col = grid1.Col
        totalamount = 0
        totaldiscount = 0
        For I = 1 To maxrow
            grid1.Row = I
            grid1.Col = 7
            totalamount = totalamount + Val(Trim(grid1.Text))
            grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.Text))
        Next
        invoicecalc
        Me.tqu.Caption = ""
        For I = 1 To maxrow
            grid1.Col = 3
            grid1.Row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(grid1.Text))
        Next
        grid1.Row = RRR
        grid1.Col = CCC
        templost = True
        
        
Exit Function
save1:
 
MsgBox "" & err.DESCRIPTION
 
        
        
End Function
Private Sub bdated_LostFocus()
''If Trim(bdated.Text) <> Trim("__/__/____") Then
''   If Not checkdate(Trim(bdated.Text), bdated) Then
''         bdated.SetFocus
''    End If
''End If
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
                If Trim(Bookname.Text) = "" Then
                    grid1.Col = 1
                    If Trim(grid1.Text) = "" Then
                        grid1.Text = Bookname.Text
                           Bookname.SetFocus
  '********* vk
                          
                          
                          If Trim(grid1.Text) = "" And Row = 1 Then
                                 grid1.Col = 2
                                 grid1.Text = ""
                                 If Trim(grid1.Text) = "" Then
                                           
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
                grid1.Text = Bookname.Text
                '/*************************
                If RS.State = 1 Then
                    RS.close
                End If
                RS.Open "books", con, adOpenDynamic, adLockReadOnly, adCmdTable
                Row = grid1.Row
                Col = grid1.Col
                If Trim(grid1.Text) <> "" Then
                    If Not RS.BOF Then
                        RS.MoveFirst
                        RS.Find "bookname='" + Trim(grid1.Text) + "'"
                        If RS.EOF Then
                            Bookname.Visible = True
                            Bookname.SetFocus
                            RS.close
                            Exit Sub
                        Else
                            
                            grid1.Col = 1
                            grid1.Text = RS(0)
                            grid1.Col = 2
                            grid1.Text = RS(1)
                        '    If Not edit Then
                                 grid1.Col = 3
                                If Trim(grid1.Text) = "" Then
                                        grid1.Text = 0
                                End If
                                q = Val(grid1.Text)
                                grid1.Col = 5
                                grid1.Text = Format(RS(3), "0.00")
                                r = RS(3)
                                '/******************
                           '------------------------------------------
                         category = returnCategory(Trim(RS(2)))
                         If Optioncash.value = True Then
                            
                                If category = "C1" Then
                                      Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                ElseIf category = "C2" Then
                                      Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                ElseIf category = "C3" Then
                                      Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII1.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                End If
                                   
                         Else
                            
                            If category = "C1" Then
                               Set kk = con.Execute("select DISCATEGORY from sledger where " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
                            ElseIf category = "C2" Then
                               Set kk = con.Execute("select CATEGORY2 from sledger where " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
                            ElseIf category = "C3" Then
                               Set kk = con.Execute("select CATEGORY3 from sledger where " & stringyear & " and subledger='" + Trim(customercode.Text) + "'")
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
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cmbdiscountcat.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                
                                ElseIf category = "C2" Then
                                
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                  
                                    End If
                                
                                ElseIf category = "C3" Then
                                
                                    If Optioncash.value = True Then
                                        Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + Trim(cboCatII1.Text) + "' and groupcode='" + Trim(RS(2)) + "'")
                                    Else
                                         Set kk = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + tempstr + "' and groupcode='" + Trim(RS(2)) + "'")
                                    End If
                                
                                
                                End If
  
                                    
                                    grid1.Col = 4
                                    If kk.BOF Then
                                        GoTo abc
                                    End If
                                    grid1.Text = Format(kk(0), "0.00")
                                    grid1.Col = 6
                                    grid1.Text = Format(kk(0), "0.00")
                                    D = kk(0)
                                Else
abc:
                                    grid1.Col = 4
                                    grid1.Text = Format(RS(4), "0.00")
                                    grid1.Col = 6
                                    grid1.Text = Format(RS(4), "0.00")
                                    D = RS(4)
                                End If
                                grid1.Col = 7
                                grid1.Text = Round(q * r, 2)
                                grid1.Col = 8
                                grid1.Text = Round((q * r) * (D / 100), 2)
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
            totalamount = totalamount + Val(Trim(grid1.Text))
            grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.Text))
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
        If Trim(Me.customercode.Text) <> "" Then
            Me.grid1.Col = 1
            Me.grid1.Row = 1
            Me.grid1.SetFocus
            Me.Grid1_Click
        Else
            If I_NO.Enabled = True Then I_NO.SetFocus
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
  rs1.Open "select *  from AgentMaster where AgentName='" & cmbAgentName.Text & "' and " & stringyear & " order by agentname", con, adOpenDynamic, adLockReadOnly, adCmdText
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

Private Sub cmdRet_Click()

DSNNew

cr.Reset

cr.ReportFileName = rptPath & "\AgentWiseCashReturn.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.WindowState = crptMaximized
cr.WindowShowPrintSetupBtn = True
cr.WindowShowPrintBtn = True
cr.Action = 1

End Sub

Private Sub cmdSalep_Click()

Screen.MousePointer = vbHourglass

Dim s As String
Dim I As Integer
Dim sum As Double
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

sum = 0
I = 1
s = ""
con.Execute "delete from CounterSale"
con.Execute "INSERT INTO countersale([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],Agent)  SELECT '0','0','0','0','0','0','0','0','0','0','0','0',Agentname from AgentWiseSale group by Agentname"

If rs2.State = 1 Then rs2.close
rs2.Open "select Agent from countersale group by Agent", con
While rs2.EOF = False
For J = 1 To 12

If RS.State = 1 Then RS.close
RS.Open "select * from BookGp where Id=" & J & "", con
If RS.EOF = False Then
I = RS("id")
s = ""
sum = 0

While RS.EOF = False
If s = "" Then
s = RS.Fields("GroupCode").value
Else
s = s & "+" & RS.Fields("GroupCode").value
End If


Set rs1 = New ADODB.Recordset
rs1.Open "select sum(netamount) from AgentWiseSale where (groupcode='" & RS.Fields("GroupCode").value & "' and agentname='" & rs2("agent") & "')", con
If Not IsNull(rs1(0)) Then
sum = sum + rs1(0)
End If
DoEvents
DoEvents

RS.MoveNext
Wend

If s = "QB" Then
 Set rs1 = New ADODB.Recordset
 rs1.Open "SELECT SUM(AMOUNT) FROM DewaliAmt where (TEXT='" & "DIWALI SPECIAL" & "' and agentname='" & rs2("agent") & "')", con
 If Not IsNull(rs1(0)) Then
 sum = sum - rs1(0)
 End If
End If



con.Execute "update CounterSale set " & "[" & I & "]" & " = " & sum & "" & " where  " & stringyear & " and agent='" & rs2(0) & "' and " & I & " = " & I

DoEvents
DoEvents

End If



Next

rs2.MoveNext
Wend

con.Execute "update CounterSale set L_link='1'"
con.Execute "update CounterSale_Head set L_link='1'"

DoEvents
DoEvents
DoEvents
DoEvents
DoEvents

DSNNew

If MsgBox("Want to show ?", vbQuestion + vbYesNo) = vbYes Then
cr.Reset
cr.ReportFileName = rptPath & "\AgentWiseCash.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.Formulas(0) = "PARTYNAME='" & "AGENT WISE SALES" & "'"
cr.WindowState = crptMaximized
cr.WindowShowPrintSetupBtn = True
cr.WindowShowPrintBtn = True
cr.WindowShowRefreshBtn = True
cr.Action = 1
End If

Screen.MousePointer = vbDefault


End Sub

Private Sub Combosldistrictcode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then txtMark.SetFocus
End Sub



Private Sub Combosldistrictcode_LostFocus()
If Combosldistrictcode.Text = "" Then
   Combosldistrictcode.SetFocus
   Exit Sub
End If
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

If Combosldistrictcode.Text <> "" Then
   rs1.Open "Select * from Districts where " & stringyear & " and Districtname = '" & Combosldistrictcode.Text & "'", con, adOpenStatic, adLockReadOnly
   If rs1.RecordCount <= 0 Then
      MsgBox "Please Select valid district.."
      Combosldistrictcode.SetFocus
   End If
End If
Set rs1 = New ADODB.Recordset
If Combosldistrictcode.Text <> "" And addmode = True Then
   rs1.Open "Select * from Districts where " & stringyear & " and Districtname = '" & Combosldistrictcode.Text & "'", con, adOpenStatic, adLockReadOnly
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

    printch = "casha_basilRet"
    ino = I_NO
    printch1 = "INVOICENO"


Printheader = False
printinvoice
SetButton Commandadd, Commandedit, Commandsave, Commanddelete
End Sub

Private Sub Commandabandon_Click()
invoiceabandon
'Me.Commandall.Enabled = False
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
    If con.Execute("Select max(invoiceno) from casha_basilRet")(0) >= Val(Trim(Me.I_NO.Text)) Then
        Me.I_NO.Text = con.Execute("Select max(invoiceno) from casha_basilRet")(0) + 1
         RS.Open "tempCash", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
         If RS.BOF Then
             RS.AddNew
         End If
         Me.I_NO.Text = RS!In + 1
         RS!In = Val(Me.I_NO.Text)
         RS.update
         RS.close
     End If
    End If
    Dim ctl As Control
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton Then
           ctl.Enabled = True
        End If
    Next
    txtMark.ListIndex = -1
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
    Me.Optioncash.SetFocus
    cboRet.ListIndex = 0
    txtMark.ListIndex = 0
End Sub
Private Sub Commandall_Click()
''Dim RS As ADODB.Recordset
''Set RS = New ADODB.Recordset
''Dim myvalue As String
''
''If Trim(Me.customercode.Text) = "" Then
''    MsgBox "Please Fill the customer detail "
''    Exit Sub
''End If
''
''myvalue = InputBox("Please enter the quantity ", "Enter the quantity: ", "1")
''
''If Len(myvalue) > 0 And Val(myvalue) > 0 Then
''
''
''
''    Grid1.Rows = 1
''    Grid1.Rows = 2
''    Grid1.col = 1
''    Grid1.row = 1
''    If RS.State = 1 Then
''        RS.Close
''    End If
''    RS.Open "select * from books order by BOOKCODE", con, adOpenDynamic, adLockReadOnly, adCmdText
''    row = Grid1.row
''    col = Grid1.col
''    If Not RS.BOF Then
''        RS.MoveFirst
''        Do While Not RS.EOF
''            Grid1.col = 1
''            Grid1.Text = RS(0)
''            Grid1.col = 2
''            Grid1.Text = RS(1)
''            Grid1.col = 3
''            If Trim(Grid1.Text) = "" Then
''                Grid1.Text = Val(myvalue)
''            End If
''            q = Val(Grid1.Text)
''            Grid1.col = 5
''            Grid1.Text = Format(RS(3), "0.00")            'rs(3)
''            r = RS(3)
''            '/******************
''            Set kk = con.Execute("select DISCATEGORY from sledger where subledger='" + Trim(customercode.Text) + "'")
''            Grid1.col = 6
''            If Trim(kk(0)) <> "" Then
''                tempstr = Trim(kk(0))
''                kk.Close
''                Set kk = con.Execute("select discountrate from DISCCATS where categorycode ='" + Trim(tempstr) + "' and groupcode='" + Trim(RS(2)) + "'")
''                Grid1.col = 4
''                If kk.BOF Then
''                    GoTo abc
''                End If
''                Grid1.Text = Format(kk(0), "0.00")
''                Grid1.col = 6
''                Grid1.Text = Format(kk(0), "0.00")
''                D = kk(0)
''            Else
''abc:
''                Grid1.col = 4
''                Grid1.Text = Format(RS(4), "0.00")
''                Grid1.col = 6
''                Grid1.Text = Format(RS(4), "0.00")
''                D = RS(4)
''            End If
''            Grid1.col = 7
''            Grid1.Text = Format(Round(q * r, 2), "0.00")
''            Grid1.col = 8
''            Grid1.Text = Format(Round((q * r) * (D / 100), 2), "0.00")
''            If Not RS.EOF Then
''                Grid1.Rows = Grid1.Rows + 1
''                Grid1.row = Grid1.row + 1
''                RS.MoveNext
''            End If
''        Loop
''        '/**fghfghgh
''        '    Grid1.col = col
''    End If
''    RS.Close
''   ' row = Grid1.row
''   ' col = Grid1.col
''    totalamount = 0
''    totaldiscount = 0
''    Me.tqu.Caption = ""
''    For I = 1 To Grid1.Rows - 1
''            Grid1.row = I
''            Grid1.col = 7
''            totalamount = totalamount + Val(Trim(Grid1.Text))
''            Grid1.col = 8
''            totaldiscount = totaldiscount + Val(Trim(Grid1.Text))
''            Grid1.col = 3
''            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(Grid1.Text))
''     Next
''     maxrow = Grid1.Rows - 1
''Else
'''Grid1_Click
''Exit Sub
''End If
''
''invoicecalc

End Sub

Private Sub Commanddelete_Click()


    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
   
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from casha_basilRet where " & stringyear & " and invoiceno=" & I_NO.Text & "", con
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select * from casha_basilRet where " & stringyear & " and invoiceno=" & I_NO.Text & "", con
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
                con.Execute ("delete  from casha_basilRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete  from cashb_basilRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
                con.Execute ("delete  from cashc_basilRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
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
    'Commandall.Enabled = True
    Commandsave.Enabled = False
    Commanddelete.Enabled = False
    Commandsearch.Enabled = False
    CommandPrint.Enabled = False
     Commandprintnh.Enabled = False
    grid1.Enabled = True
    'Commandall.Enabled = False
    Me.customercode.Enabled = True
    Edit = True
    I_NO_LostFocus
    i_dt.Enabled = True
    i_dt.SetFocus
    ' cashc_basilTMP creation start
    DoEvents
    con.Execute "Delete  from CASHCTMP_basilRet where " & stringyear
    DoEvents
    'CON.Execute ("insert into CASHCTMP_basilRet  select * from cashc_basilRet where INVOICENO = " + Trim(I_NO.Text))
    con.Execute ("insert into CASHCTMP_basilRet(INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT," & _
    "Rate,AMOUNT,DEBITORCREDIT,[TEXT],RYN,Fyear,setupid)  select INVOICENO,INVOICEDATE,GENLEDGER," & _
    "SUBLEDGER,GAMOUNT,Rate,AMOUNT,DEBITORCREDIT,[TEXT],RYN,Fyear,setupid from cashc_basilRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
    
    
    Dim KS As Long
    KS = 1
    For L = 1 To 15000
      PP = 0
    Next L
    
    DoEvents
    On Error Resume Next
    addoredit = True
    HIT
    VB.Screen.ActiveForm.ActiveControl.SelLength = 2
    Me.Commandother.Enabled = True
End Sub
Private Sub Commandother_Click()

Commandsave.Enabled = True
searchForm = "cashbasilret"
frmEndPartTrans.Show

End Sub
Private Sub CommandPrint_Click()
   
   
    printch = "casha_basilRet"
    ino = I_NO
    printch1 = "INVOICENO"
   
   
   Printheader = True
   printinvoice
End Sub
Private Sub CommandReturn_Click()
'   Dim RS As New ADODB.Recordset
'   RS.Open "tempCASH", CON1, adOpenDynamic, adLockOptimistic, adCmdTable
'   If RS.BOF Then
'       RS.AddNew
'   End If
'   RS!In = CON.Execute("Select max(invoiceno) from casha_basilRet")(0)
'   RS.Update
'   RS.close
   Unload Me
   addoredit = False
'   'MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()
   
    
    
    
    Dim rs_h As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
   
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from casha_basilRet where " & stringyear & " and invoiceno=" & I_NO.Text & "", con
    If rs1.EOF = False Then
       
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select * from casha_basilRet where " & stringyear & " and invoiceno=" & I_NO.Text & "", con
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
    If Trim(Combosldistrictcode) = "" Then
        MsgBox "Please Enter District"
        Exit Sub
    End If
       
    
Else
   
   
    If Edit = False Then
      If (I_OB <> "" And txtMark <> "") Then
          Party_Remove_FromOrder Trim(Me.customercode.Text), txtMark, Trim(I_OB)
      End If
    End If
   
   
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
grid1.Row = 1
grid1.Col = 1
If Trim(grid1.Text) = "" Then
   MsgBox "Please Enter item.... "
   Exit Sub
End If

If Edit = False Then

    If check_Duplikate("casha_basilRet", I_NO.Text) = True Then
       MsgBox "This Inv. Number Already Exist ..", vbCritical
       Exit Sub
    End If

End If


If Trim(I_NO.Text) <> "" And Trim(i_dt.Text) <> "" And Trim(customercode.Text) <> "" Then
   If Edit Then
      con.Execute ("delete  from casha_basilRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
      con.Execute ("delete  from cashb_basilRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
      con.Execute ("delete  from cashc_basilRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
   End If
   If RS.State = 1 Then RS.close
   LAMOUNT = 0
   RS.Open "select * from casha_basilRet where  " & stringyear & " and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
   If Not Edit Then
again:
      If con.Execute("Select max(invoiceno) from casha_basilRet")(0) >= Val(Trim(Me.I_NO.Text)) Then
        ' Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
         'GoTo again
      End If
   End If
   RS.AddNew
   
   
   RS!setupid = setupid
   RS!fyear = session
   
   RS!Godown = txtMark.Text
   RS!invoiceNo = Val(Me.I_NO.Text)
   RS!invoiceDate = Me.i_dt.Text
   RS!Genledger = Trim(Me.Genledger.Text)
   RS!subledger = Trim(Me.customercode.Text)
   RS!orderby = Trim(Me.I_OB.Text)
   RS!bundles = Trim(Me.bundles)

   
   
   RS!netamount = Round(Val(Trim(Me.mna.Caption)), 2)
   RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
   RS!txt1 = Trim(frmEndPartTrans.T1TEXT.Text)
   RS!txt1a = Val(Trim(frmEndPartTrans.T1.Text))
   RS!txt2 = Trim(frmEndPartTrans.T2TEXT.Text)
   RS!txt2a = Val(Trim(frmEndPartTrans.T2.Text))
   RS!baa = Val(Trim(frmEndPartTrans.T3TEXT.Text))
   
   RS!District = Combosldistrictcode.Text
   RS!CASHPARTYNAME = textbox.Text
   RS!agentname = cmbAgentName.Text
   RS!discat = cmbdiscountcat.Text
   RS!discat2 = cboCatII.Text
   RS!discat3 = cboCatII1.Text
  
err1:
   If Not Edit Then
      If con.Execute("Select max(invoiceno) from casha_basilRet")(0) >= Val(Trim(Me.I_NO.Text)) Then
         'Me.I_NO.Text = Str(Val(Trim(Me.I_NO.Text)) + 1)
         'rs!INVOICENO = Val(Me.I_NO.Text)
         On Error GoTo err1
      End If
   End If
   RS.update
   
   
   On Error GoTo 0
   RS.close
   RS.Open "select * from cashb_basilRet where " & stringyear & " and invoiceno<=0", con, adOpenDynamic, adLockOptimistic
   Dim I As Integer
   RRRR = grid1.Row
   CCCC = grid1.Col
   For I = 1 To maxrow
       grid1.Row = I
       grid1.Col = 1
       If Trim(grid1.Text) <> "" Then
          grid1.Col = 3
          If Val(Trim(grid1.Text)) > 0 Then
             grid1.Col = 5
            If Val(Trim(grid1.Text)) > 0 Then
               RS.AddNew
               grid1.Col = 1
               RS!invoiceNo = Val(Me.I_NO.Text)
               RS!invoiceDate = Me.i_dt.Text
               RS!Genledger = Trim(Me.Genledger.Text)
               RS!subledger = Trim(Me.customercode.Text)
               RS!Bookcode = Trim(grid1.Text)
               grid1.Col = 3
               RS!QUANTITY = Trim(grid1.Text)
               grid1.Col = 5
               RS!rate = Trim(grid1.Text)
               grid1.Col = 7
               RS!amount = Trim(grid1.Text)
               LAMOUNT = Val(Trim(grid1.Text))
               grid1.Col = 4
               RS!PRINTORDER = Trim(grid1.Text)
               grid1.Col = 6
               RS!discount = Trim(grid1.Text)
               grid1.Col = 8
               RS!netamount = LAMOUNT - Trim(grid1.Text)
               LAMOUNT = 0
               RS!agentname = Trim(Me.cmbAgentName.Text)
               RS!setupid = setupid
               RS!fyear = session

               RS.update
            End If
         End If
     End If
  Next
  RS.close
  grid1.TopRow = 1
  RS.Open "select * from cashc_basilRet where " & stringyear & " and invoiceno<=0 ", con, adOpenDynamic, adLockOptimistic
  '/******
  'Dim I, x As Integer
   Dim temprs As ADODB.Recordset
   Set temprs = New ADODB.Recordset
       For I = 1 To frmEndPartTrans.vs.rows - 1
           frmEndPartTrans.vs.Row = I
           frmEndPartTrans.vs.Col = 0
           If Trim(frmEndPartTrans.vs.Text) <> "" Then
              RS.AddNew
                         
              RS!setupid = setupid
              RS!fyear = session

              RS!invoiceNo = Val(Me.I_NO.Text)
              RS!invoiceDate = Me.i_dt.Text
              RS!gamount = Round((Me.totalamount - Me.totaldiscount), 2)
              RS!Text = Trim(frmEndPartTrans.vs.Text)
              If temprs.State = 1 Then temprs.close
              If Edit Then
                 temprs.Open "select * from CASHCTMP_basilRet WHERE " & stringyear & " and INVOICENO = " & Val(Me.I_NO.Text) & "", con, adOpenDynamic, adLockReadOnly, adCmdText
                 If frmEndPartTrans.vs.Text <> "" Then
                    temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.Text) + "'"
                    RS!Genledger = temprs!Genledger & ""
                    RS!subledger = temprs!subledger & ""
                    RS!DebitorCredit = temprs!DebitorCredit & ""
                    RS!RYN = temprs!RYN & ""
                End If
                temprs.close
              Else
                 temprs.Open "select * from INVOICEEND where " & stringyear & " and type='cashbasilret'", con, adOpenDynamic, adLockReadOnly, adCmdText
                 If frmEndPartTrans.vs.Text <> "" Then
                    temprs.Find "TEXT='" + Trim(frmEndPartTrans.vs.Text) + "'"
                    RS!Genledger = temprs!Genledger & ""
                    RS!subledger = temprs!subledger & ""
                    RS!DebitorCredit = temprs!DebitorCredit & ""
                    RS!RYN = temprs!RYN & ""
                 End If
                 temprs.close
              End If
              frmEndPartTrans.vs.Col = 1
              RS!rate = Val(Trim(frmEndPartTrans.vs.Text))
              If Val(Trim(frmEndPartTrans.vs.Text)) > 0 Then
                 RS!amount = Round((Me.totalamount - Me.totaldiscount), 2) * Round((Val(Trim(frmEndPartTrans.vs.Text)) / 100), 2)
              Else
                frmEndPartTrans.vs.Col = 2
                RS!amount = Val(Trim(frmEndPartTrans.vs.Text))
              End If
              
   
              RS.update
          End If
      Next
      RS.close
      
       con.Execute ("delete from CASHCTMP_basilRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text))
       
       SAVED = True
       
  End If
  If SAVED Then
      Unload frmEndPartTrans
   
      MsgBox "Record Saved"
      
      Me.customercode.Enabled = False
      Me.grid1.Enabled = False
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
   searchType = "inv"
   
   sqlQry = "select InvoiceNo,InvoiceDate,Subledger,NetAmount from casha_basilRet where InvoiceNo"
   orderby = "order by InvoiceNo"

   
   popuplist10 "select InvoiceNo,InvoiceDate,Subledger,NetAmount from casha_basilRet where " & stringyear & "  order by InvoiceNo", con
End Sub

Private Sub Commandsearch_GotFocus()
If PopUpValue1 <> "" Then
     I_NO.Text = PopUpValue1
     I_NO_LostFocus
     PopUpValue1 = ""
End If

End Sub

Private Sub customercode_GotFocus()
' SendKeys "{DOWN}"
End Sub

Private Sub customercode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

If Optioncredit.value = True Then
   txtMark.SetFocus
 End If
End If
End Sub

Private Sub customercode_KeyPress(KeyAscii As Integer)
' On Error Resume Next
'   If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
'       ' SendKeys "{tab}"
'       ' Exit Sub
'   End If
'
'    If KeyAscii = 13 Then
'       'SendKeys "{DOWN}"
'
'       'customercode.Visible = False
'       'SendKeys "{TAB}"
'       'bundles.SetFocus
'    End If
End Sub
Private Sub customercode_LostFocus()

        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        SendKeys "{DOWN}"
        RS.Open "select * from sledger where " & stringyear & " and gledger='SUNDRY DEBTORS' and subledger='" + Trim(customercode.Text) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
        If (RS.RecordCount <= 0 Or RS.EOF = True) Then
           customercode.SetFocus
           HIT
           RS.close
           Exit Sub
        End If
        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        If RS!distcode <> "" And addmode = True Then
            rs1.Open "Select * from Districts where " & stringyear & " and Districtname = '" & RS!distcode & "'", con, adOpenStatic, adLockReadOnly
            If rs1.RecordCount > 0 Then
                Me.cmbAgentName = IIf(IsNull(rs1!agentname), "", rs1!agentname)
            End If
             Combosldistrictcode.Text = RS!distcode
        Else
        Combosldistrictcode.Text = RS!distcode & ""
        End If
        Me.textbox.Text = Me.customercode.Text
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
    'SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    mna.Enabled = True
    Label2.Enabled = True
End Sub

Function returnCategory(s As String) As String
    Dim s1 As New ADODB.Recordset
    If s1.State = 1 Then s1.close
    
    s1.Open "select category from [groups] where " & stringyear & " and groupcode='" & s & "'", con
    If s1.EOF = False Then
       returnCategory = s1(0)
    End If
    
End Function



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If grid1.Row >= 1 Then
           grid1.RemoveItem grid1.Row
           a = grid1.Text
           tempmeb.Text = a
           a = templost
           grid1.SetFocus
          End If
   End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If UCase(Trim(VB.Screen.ActiveControl.Name)) = UCase(Trim("CUSTOMERCODE")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("txtmark")) Then
        If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
             SendKeys "{tab}"
             Exit Sub
        End If
          If addmode = True Then
                SendKeys "{DOWN}"
           End If
            SendKeys "{TAB}"
        Else
            If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("cboret")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("grid1")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("tempmeb")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bookname")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("bundles")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("txtmark")) Then
                SendKeys ("{TAB}")
            End If
        End If
    End If
End Sub


Private Sub Form_Load()
'On Error Resume Next

Me.Left = 100
Me.Top = 100
Me.Width = 10900
Me.Height = 8400

Me.grid1.Left = 150


Screen.MousePointer = vbHourglass

   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset


    txtMark.Clear

    If RS.State = 1 Then RS.close
    RS.Open "select * from GodownMaster order by id", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.txtMark.AddItem RS(0)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close





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
    Me.Top = 0
    Me.Left = 0

    grid1.rows = 2
    grid1.Cols = 1
    grid1.rows = 10
    grid1.Cols = 9
    grid1.Row = 0
    grid1.Col = 1
    grid1.Text = "Book Code "
    grid1.Col = grid1.Col + 1
    grid1.Text = "Book Name"
    grid1.Col = grid1.Col + 1
    grid1.Text = "Quantity"
    grid1.Col = grid1.Col + 1
    grid1.Text = "Print. Ord."
    grid1.Col = grid1.Col + 1
    grid1.Text = "Rate"
    grid1.Col = grid1.Col + 1
    grid1.Text = "Disc %"
    grid1.Col = grid1.Col + 1
    grid1.Text = "Amount"
    grid1.Col = grid1.Col + 1
    grid1.Text = "Disc. Amount"
    grid1.RowHeight(0) = grid1.CellHeight + 50
    grid1.ColWidth(0) = 150
    grid1.ColWidth(1) = 1100
    grid1.ColWidth(2) = 2000
    grid1.ColWidth(3) = 750
    grid1.ColWidth(4) = 750
    grid1.ColWidth(5) = 1200
    grid1.ColWidth(6) = 800
    grid1.ColWidth(7) = 1200
    grid1.ColWidth(8) = 1200
    Me.CommandPrint.Enabled = True
    Me.Commandprintnh.Enabled = True
    
If RS.State = 1 Then RS.close
RS.Open "select Distinct categorycode from DISCCATS order by categorycode", con, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Do While Not RS.EOF
        Me.cmbdiscountcat.AddItem RS!categorycode
        cboCatII.AddItem RS!categorycode
        cboCatII1.AddItem RS!categorycode
        If Not RS.EOF Then
        RS.MoveNext
        End If
    Loop
End If
RS.close

'============================================================================
'----------------------------------------------------------------
Set RS = con.Execute("exec BookQry '" & session & "'," & main.setupid & "")
If Not RS.BOF Then
    Do While Not RS.EOF
        Me.Bookcode.AddItem RS(1)
        Me.Bookname.AddItem RS(0)
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If
    
Genledger.Text = "SUNDRY DEBTORS"
Set RS = con.Execute("exec fatch_ledger '" & Genledger.Text & "','" & session & "'," & main.setupid & "")

If Not RS.BOF Then
    Do While Not RS.EOF
        Me.customercode.AddItem RS(0)
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If

'=============================================================================
'*******Agent  combo fill
If RS.State = 1 Then RS.close
RS.Open "select  Agentname from AgentMaster where " & stringyear & " order by agentname", CON_blue, adOpenForwardOnly, adLockReadOnly
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





    Bookcode.Left = grid1.Left
    Bookcode.Visible = False
    Bookname.Visible = False
    grid1.rows = 100
    For I = 1 To 99
        grid1.RowHeight(I) = 300
    Next
    Bookcode.Width = 1230
    Bookname.Width = 3830
    amount.Width = rate.Width

       kk.Open "SELECT MAX(INVOICENO) FROM casha_basilRet where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
       If Not IsNull(kk(0)) Then
          Me.I_NO.Text = Trim(Str(kk(0)))
          I_NO_LostFocus
       Else
          Me.I_NO.Text = "1"
       End If
       kk.close

   mna.Enabled = True
   Label2.Enabled = True
   Commanddelete.Enabled = True
   Commandedit.Enabled = True
   Commandsave.Enabled = False
   lastrow = 0
   lastcol = 1
   
 
    Picture5.Enabled = True
    If RS.State = 1 Then RS.close
    RS.Open "select * from DISTRICTS where " & stringyear & " order by DISTRICTNAME", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS!DISTRICTNAME
            Me.cmbareaname.AddItem RS!DISTRICTNAME
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If


    
    SetButton Commandadd, Commandedit, Commandsave, Commanddelete
    cmdSalep.Enabled = True
    BackColorFrom Me, 1


Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_Resize()

panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub freight_LostFocus()
freight = UCase(freight)
End Sub

 Sub Grid1_Click()
If Trim(Me.customercode.Text) <> "" Then
Dim PREVROW As Integer
Dim prevcol As Integer
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
prevcol = Me.grid1.Col
PREVROW = Me.grid1.Row

If Me.grid1.Row > 1 Then
    grid1.Row = grid1.Row - 1
    grid1.Col = 1
    If Trim(grid1.Text) <> "" Then
        grid1.Row = PREVROW
        grid1.Col = prevcol
        If Trim(Me.customercode.Text) <> "" Then
            If Me.customercode.Enabled = True Then
                Me.customercode.Enabled = False
            End If
            grid1.Col = 1
            If prevcol > 1 And Trim(grid1.Text) = "" Then
                grid1.Col = 2
                SendKeys Chr(13)
            Else
                grid1.Col = prevcol
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
        Me.grid1.Col = 1
        If prevcol > 1 And Trim(grid1.Text) = "" Then
            Me.grid1.Col = 2
            Me.grid1.SetFocus
            SendKeys Chr(13)
        Else
        'IF GRID1.COL
            Me.grid1.Col = prevcol
            Me.grid1.SetFocus
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
            
            Select Case grid1.Col
            Case 1, 3, 4, 5, 6
                Bookname.Visible = False
                tempmeb.Visible = True: tempmeb.Enabled = True
                tempmeb.ZOrder
                tempmeb.Width = grid1.ColWidth(grid1.Col)
                tempmeb.Left = grid1.CellLeft + leftAlign
                tempmeb.Top = grid1.Top + grid1.CellTop '- 50
 
                If grid1.Col <> 1 Then
                    If grid1.Col <> 3 Then
                        tempmeb.Text = Format(grid1.Text, "0.00")
                        
                    Else
                        tempmeb.Text = Format(grid1.Text, "0")
                    End If
                   
                Else
                    tempmeb.Text = grid1.Text
                End If
            Case 2
                tempmeb.Visible = False
                Bookname.Visible = True: Bookname.Enabled = True
                Bookname.ZOrder
                Bookname.Text = grid1.Text
                Bookname.Top = grid1.Top + grid1.CellTop
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
                    SendKeys Chr(KeyAscii)
                End If
            Case 1, 3, 4, 5, 6
                mprevcol = grid1.Col
                tempmeb.SetFocus
            Case Else
                If KeyAscii = 13 Then
                    SendKeys "{RIGHT}"
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
   PopupMenu dd, , grid1.Left + X, grid1.Top + Y
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
        MsgBox "Enter Valid Date !!", vbInformation
        i_dt.SetFocus
        Exit Sub
    End If
End If

End Sub
Private Sub I_DTOB_LostFocus()
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
        MsgBox "Estimate No cannot be null"
        'I_NO.SetFocus
    Else
        If RS.State = 1 Then RS.close
        RS.Open "Select * from  casha_basilRet where " & stringyear & " and INVOICENO = " + Trim(I_NO.Text) + "", con, adOpenStatic, adLockReadOnly
        If RS.EOF Then
            If addoredit = False Then
                 MsgBox "Estimate No not found"
                 SetButton Commandadd, Commandedit, Commandsave, Commanddelete
                 Exit Sub
            End If
            Exit Sub
        End If
        If addoredit Then
            MsgBox "Estimate No already exist..."
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
        
        I_NO.Text = RS!invoiceNo
        cboRet.Text = RS!Sale_Return
        Me.i_dt.Text = RS!invoiceDate
        txtMark = UCase(Trim(RS!Godown))
        Me.Genledger.Text = Trim(RS!Genledger)
        Me.customercode.Text = Trim(RS!subledger)
        Me.textbox.Text = Trim(RS!subledger)
        Me.I_OB.Text = Trim(RS!orderby)
        'Me.I_DTOB.Text = IIf(IsNull(RS!ORDERDATE), "__/__/____", RS!ORDERDATE)
        'Me.marka.Text = Trim(rs!marka)
        Me.bundles = Trim(RS!bundles)
        'Me.through.Text = rs!through
        'Me.through1.Text = rs!through1
        'Me.station.Text = RS!station
        'Me.biltno.Text = Trim(RS!biltyno)
        'Me.bdated = IIf(IsNull(RS!BILTYDATE), "__/__/____", RS!BILTYDATE)
        'Me.freight = Trim(RS!freight)
        'Me.weight = Trim(rs!weight)
        'Me.labelbybank = Format(Round(Val(RS!baa), 2), "0.00")
        mna.Caption = Format(Round(Val(RS!netamount), 2), "0.00")
        'Me.cmbtransportname.Text = IIf(IsNull(RS!transportname), "", RS!transportname)

      
        
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
        cboCatII.Text = IIf(IsNull(RS!discat2), "", RS!discat2)
        cboCatII1.Text = IIf(IsNull(RS!discat3), "", RS!discat3)
        
        RS.close
        grid1.TopRow = 1
    '*/**/*/*/*/*//*/*
    If RS.State = 1 Then RS.close
    RS.Open "Select * from cashb_basilRet where " & stringyear & " and INVOICENO =" + Trim(I_NO.Text) + "  order by SNO ", con, adOpenStatic, adLockReadOnly
    If Not RS.EOF Then
            grid1.Row = 1
            grid1.Col = 1
            Do While Not RS.EOF
               If Trim(RS!invoiceNo) = Trim(I_NO.Text) Then
                grid1.Col = 1
                grid1.Text = Trim(RS!Bookcode)
                If kk.State = 1 Then
                    kk.close
                End If
                kk.Open "select * from books where " & stringyear & " and bookcode='" + Trim(RS!Bookcode) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
                grid1.Col = 2
                grid1.Text = Trim(kk!Bookname)
                grid1.Col = 3
                grid1.Text = Trim(RS!QUANTITY)
                grid1.Col = 5
                grid1.Text = Format(Round(RS!rate, 2), "0.00")
                grid1.Col = 7
                grid1.Text = Format(Round(RS!amount, 2), "0.00")
                grid1.Col = 4
                grid1.Text = Format(Round(RS!PRINTORDER, 2), "0.00")
                grid1.Col = 6
                grid1.Text = Format(Round(RS!discount, 2), "0.00")
                grid1.Col = 8
                grid1.Text = Format(Round(RS!amount * (RS!discount / 100), 2), "0.00")
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
            totalamount = totalamount + Val(Trim(grid1.Text))
            grid1.Col = 8
            totaldiscount = totaldiscount + Val(Trim(grid1.Text))
        Next
        mga.Caption = Format(Round(totalamount, 2), "0.00")
        mgd.Caption = Format(Round(totaldiscount, 2), "0.00")
        Me.tqu.Caption = ""
        For I = 1 To maxrow
            grid1.Col = 3
            grid1.Row = I
            Me.tqu.Caption = Val(Trim(Me.tqu.Caption)) + Val(Trim(grid1.Text))
        Next
        grid1.Row = RRR
        grid1.Col = CCC
    End If
    mna.Enabled = True
    Label2.Enabled = True
    'Me.Commandother.Enabled = True
    Me.Commandother.Enabled = False
    
End Sub

Private Sub I_OB_GotFocus()
Dim trs As New ADODB.Recordset
trs.Open " SELECT DISTCODE    FROM SLEDGER  WHERE " & stringyear & " and SUBLEDGER='" & customercode.Text & "'", con, adOpenStatic, adLockOptimistic, adCmdText
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
       cmbdiscountcat.Visible = True
       lbldis.Visible = True
       
       Combosldistrictcode.Visible = True
       lbldis1.Visible = True
       cboCatII.Visible = True
       
       LBLDIS2.Visible = True
       cboCatII1.Visible = True

       
 End If


End Sub

Private Sub Optioncredit_Click()
If Optioncredit.value = True Then
       Label4.Visible = False
       Combosldistrictcode.Visible = False
       lbldis.Visible = False
       cmbdiscountcat.Visible = False
       
       lbldis1.Visible = False
       cboCatII.Visible = False
       
       LBLDIS2.Visible = False
       cboCatII1.Visible = False

  End If

End Sub

Private Sub station_LostFocus()
station = UCase(station)
End Sub

Private Sub tempmeb_Change()
If grid1.Col = 1 Or grid1.Col = 2 Then
    grid1.Text = tempmeb.Text
Else
    If grid1.Col = 3 Then
        grid1.Text = Format(tempmeb.Text, "0")
    Else
        grid1.Text = Format(tempmeb.Text, "0.00")
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
                    '    RS.Close
                    'End If
                    'RS.Open "books", CON, adOpenDynamic, adLockReadOnly, adCmdTable
                    Set RS = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & Trim(grid1.Text) & "'")
                    If Not RS.BOF Then
                        'RS.MoveFirst
                        'RS.Find "bookcode='" + Trim(Grid1.Text) + "'"
                        If RS.EOF And Trim(grid1.Text) <> "" Then
                            RS.close
                            Exit Sub
                        Else
                            RS.close
                        If Trim(grid1.Text) <> "" Then
                                grid1.Col = 3
                            Else
                                grid1.Col = 2
                            End If
                        End If
                    Else
                        If Trim(grid1.Text) <> "" Then
                            grid1.Col = 3
                        Else
                            grid1.Col = 2
                        End If
                    End If
                    grid1.SetFocus
                    Grid1_Click
                
                Case 3
                    If Val(tempmeb.Text) > 0 Then
                        grid1.Col = grid1.Col + 2
                        grid1.SetFocus
                        Grid1_Click
                    End If
                Case 4
                    grid1.Col = grid1.Col + 2
              '       SendKeys "{LEFT}"
                    grid1.SetFocus
                    Grid1_Click
                Case 5
                    If Val(tempmeb.Text) > 0 Then
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

'Private Sub textbox_KeyDown(KeyCode As Integer, Shift As Integer)
' If Optioncredit.Value = True Then
'    'bundles.SetFocus
'    'txtMark.SetFocus
' End If
'End Sub

Private Sub textbox_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Then
   'cmbareaname.Visible = True
   Me.customercode.Enabled = True
   Me.customercode.Visible = False
  'Me.customercode.Height = 1100
   Me.cmbareaname.ZOrder
  'Me.cmbareaname.SetFocus
   


   
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
          kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
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
    rs1.Open "casha_basilRet", con, adOpenDynamic, adLockReadOnly, adCmdTable
    rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
   
    Line = Line + 1
    If Not rs1.EOF Then
        Print #1, " To,"; Tab(T1 - 3); rs1!subledger; Tab(T5); "Estimate No. : "; Trim(rs1!invoiceNo); Tab(T8 + 5); "Dt. "; Tab(T8 + 12); rs1!invoiceDate 'Chr(27) + Chr(15);
            If kkk.State = 1 Then
                kkk.close
            End If
            kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
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
            kk.Open "select * from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " order by discount,printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                        tdata.Open "select sum(amount) from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
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
           kk.Open "Select * from cashc_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
           kk.Open "Select * from casha_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
       
       
            'Print #1, Tab(0); repli("-", 120)
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            Dim LEFTM As Integer
            LEFTM = 5
            CNSetup
            tempdata.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
            'Print #1, Tab(1); "E.& O.E"
            'Print #1, Tab(LEFTM); tempdata!COURT; Tab(LEFTM + (paperWidth - ((Len(tempdata!COURT) + Len(tempdata!CNAME)) * 0.75))); "FOR " + Trim(tempdata!CNAME)
            Print #1, Tab(0); repli("-", 120)
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
kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
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
Print #1, Tab(0); Chr(27) + Chr(15) + Chr(14); Tab(30); "Estimate"; Chr(20)
Print #1, Tab(0); repli("*", 148)
Line = Line + 3
If rs1.State = 1 Then
   rs1.close
End If

If rs1.State = 1 Then
    rs1.close
End If
rs1.Open "casha_basilRet", con, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
If Not rs1.EOF Then
    Print #1, " To,"; Tab(T1 - 3); IIf(Optioncash.value = True, rs1!CASHPARTYNAME, rs1!subledger); Tab(T5); "Estimate No. : "; Trim(rs1!invoiceNo); Tab(T8 + 5); "Dt. :"; Tab(T8 + 12); rs1!invoiceDate 'Chr(27) + Chr(15);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
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
    kk.Open "select * from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " order by discount,printorder", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                tdata.Open "select sum(amount) from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
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
        kk.Open "Select * from cashc_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
                        Print #1, Tab(T5); Trim(kk!Text) + "    " + Trim(Format(Str(Round(kk!rate, 2)), "0.00")); Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
                    Else
                        Print #1, Tab(T5); Trim(kk!Text); Tab(T8 + 5); rsets(Trim(Format(Str(Round(kk!amount, 2)), "0.00")), 12)
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
        kk.Open "Select * from casha_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
        tempdata.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
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
    kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
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
        'Print #1, Tab(1); "E.& O.E"
        'Print #1, Tab(1); kkk!COURT; Tab(75); "FOR " + Trim(kkk!CNAME)
        Print #1, ""
        'Print #1, Tab(1); Chr(27) + Chr(71); "Continued on Page : " & Pno; Chr(27) + Chr(72)
        Print #1, Tab(0); repli("-", 96)
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
''Print #1, Chr(27) + Chr(71); IIf(FooterYes = True, "Continued from Page : " & Pno - 1, ""); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("Estimate")))) / 2 - 10); Chr(14); "***ESTIMATE SALES***"; Chr(20); Tab(50); IIf(Printheader = True, kkk!uptt, "")
''Line = Line + 1
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
rs1.Open "casha_basilRet", con, adOpenDynamic, adLockReadOnly, adCmdTable
rs1.Find "invoiceno='" + Trim(Me.I_NO.Text) + "'"
'Tab(20); Mid$(rs1!SUBLEDGER, 1, 5);
If Not rs1.EOF Then
    Print #1, Chr(27) + Chr(71); " To,"; Tab(7); IIf(Optioncash.value = True, "", Mid$(rs1!subledger, 1, 5)); Tab(48); "Estimate No. : "; Trim(rs1!invoiceNo); Tab(82); "Dt. : "; rs1!invoiceDate; Chr(27) + Chr(72);
    Line = Line + 1
    If kkk.State = 1 Then
        kkk.close
    End If
    kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs1!subledger) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
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
    kk.Open "select * from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " order by discount,sno ", con, adOpenDynamic, adLockReadOnly, adCmdText
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
                'Print #1, rsets(Trim(Str(sno)), 4); Tab(11); Trim(tdata!Bookname); Tab(T5 - 4); rsets(Trim(Str(kk!quantity)), 5); Tab(T6 + 1); rsets(Trim(Format(Str(kk!rate), "0.00")), 8); Tab(T7 - 1); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
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
                tdata.Open "select sum(amount)as sumamt, sum(amount-netamount) as sumdis from cashb_basilRet where " & stringyear & " and invoiceno=" + Trim(rs1!invoiceNo) + " and discount=" + Trim(Str(cdiscount)) + " group by discount", con, adOpenDynamic, adLockReadOnly, adCmdText
                If Not tdata.BOF Then
                   Print #1, Tab(68); rsets(Trim(Format(Str(tdata(0)), "0.00")), 12)
                   'Print #1, Tab(30); "Less Discount @ " + Trim(Format(Str(cdiscount), "0.00")) + " %"; Tab(68); rsets(Trim(Format(Str(tdata(0) * cdiscount / 100), "0.00")), 12); Tab(84); rsets(Trim(Format(Str(tdata(0) - (tdata(0) * cdiscount / 100)), "0.00")), 12)
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
    kk.Open "Select * from cashc_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not kk.BOF Then
       Do While Not kk.EOF
             If kk!amount > 0 Then
                    If Trim(kk!DebitorCredit) = Trim("Credit") Then
                        netamount = netamount + kk!amount
                    Else
                        netamount = netamount - kk!amount
                    End If
                    If kk!rate > 0 Then
                        Print #1, Tab(60); Trim(kk!Text) + "    " + Trim(Format(Str(kk!rate), "0.00")); Tab(84); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
                    Else
                        Print #1, Tab(60); Trim(kk!Text); Tab(84); rsets(Trim(Format(Str(kk!amount), "0.00")), 12)
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
        kk.Open "Select * from casha_basilRet where " & stringyear & " and invoiceno=" + Trim(Me.I_NO.Text), con, adOpenDynamic, adLockReadOnly, adCmdText
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
        
        Dim tempdata As ADODB.Recordset
        Set tempdata = New ADODB.Recordset
        CNSetup
        tempdata.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
        'Print #1, Tab(1); "E.& O.E"
        'Print #1, Tab(1); tempdata!COURT; Tab(65); "FOR " + Trim(tempdata!CNAME)
        Print #1, Tab(0); repli("-", 96)
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
Private Sub txtMark_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
         SendKeys "{tab}"
      End If
End Sub

