VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAgreement 
   Caption         =   "Customer Agreement"
   ClientHeight    =   10692
   ClientLeft      =   60
   ClientTop       =   396
   ClientWidth     =   16920
   Icon            =   "frmAgreement.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10692
   ScaleWidth      =   16920
   Begin TabDlg.SSTab SSTab1 
      Height          =   10224
      Left            =   72
      TabIndex        =   6
      Top             =   12
      Width           =   16800
      _ExtentX        =   29633
      _ExtentY        =   18034
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   8454016
      TabCaption(0)   =   "Page - 1"
      TabPicture(0)   =   "frmAgreement.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(11)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(12)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(13)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3(14)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(15)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3(16)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3(17)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3(18)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3(19)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3(20)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label3(21)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3(22)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3(23)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3(24)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label3(41)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "vs1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "VSExamtion"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "vs"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtDate"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtName"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtAgmNo"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtPName"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtAddress1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtAddress2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtAddress3"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtAddress4"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtMobile"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtEmail"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtSubject"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtSub_yrs"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtdearsir"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtwarm1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtwarm2"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtwarm3"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtarea"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtexpSale"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtSpNoteA"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtSpNoteB"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtSpNoteC"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtSession1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdAdd_1"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmdSave_2"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "CommandPrint"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cr"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "CommandReturn"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Commanddelete"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtSPNoteForExamation"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Check1_OldNew"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Commandedit"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Frame1"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Check1_TOD"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Check2_CD"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Check3_BaseDis"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "cmdRepQty"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).ControlCount=   56
      TabCaption(1)   =   "Page - 2"
      TabPicture(1)   =   "frmAgreement.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo1_State"
      Tab(1).Control(1)=   "txtCashDepRem1"
      Tab(1).Control(2)=   "txtTurnOverSp"
      Tab(1).Control(3)=   "txtTurnOverSp1"
      Tab(1).Control(4)=   "txtTransportation1"
      Tab(1).Control(5)=   "txtTurnOverSp3"
      Tab(1).Control(6)=   "Check1_TDExamte"
      Tab(1).Control(7)=   "txtCashDepositAlpha"
      Tab(1).Control(8)=   "txtCashDepositSch_Heading"
      Tab(1).Control(9)=   "txtTrans_Alfabat"
      Tab(1).Control(10)=   "txtTransportation"
      Tab(1).Control(11)=   "MSFlexGrid1"
      Tab(1).Control(12)=   "txtCashDis_Alpha"
      Tab(1).Control(13)=   "txtTurnDis_Alpha"
      Tab(1).Control(14)=   "txtCashMinExtra1"
      Tab(1).Control(15)=   "txtCashMinExtra2"
      Tab(1).Control(16)=   "txtCashMinAmt2"
      Tab(1).Control(17)=   "txtCashMinAmt1"
      Tab(1).Control(18)=   "Check1_cashdep"
      Tab(1).Control(19)=   "txtCashDepSpNot"
      Tab(1).Control(20)=   "txtCashDepoSch"
      Tab(1).Control(21)=   "Command3"
      Tab(1).Control(22)=   "Check1_turnOver"
      Tab(1).Control(23)=   "txtTurnOverDis"
      Tab(1).Control(24)=   "vs_page2"
      Tab(1).Control(25)=   "vsCashDeposit"
      Tab(1).Control(26)=   "txtCashDepositSch"
      Tab(1).Control(27)=   "Label3(38)"
      Tab(1).Control(28)=   "Label3(31)"
      Tab(1).Control(29)=   "Label3(30)"
      Tab(1).Control(30)=   "Label3(27)"
      Tab(1).Control(31)=   "Label3(26)"
      Tab(1).Control(32)=   "Label3(25)"
      Tab(1).Control(33)=   "Label3(9)"
      Tab(1).Control(34)=   "Label3(2)"
      Tab(1).Control(35)=   "Label3(0)"
      Tab(1).ControlCount=   36
      TabCaption(2)   =   "Page - 3"
      TabPicture(2)   =   "frmAgreement.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "Label3(42)"
      Tab(2).Control(2)=   "Label3(44)"
      Tab(2).Control(3)=   "Label3(45)"
      Tab(2).Control(4)=   "Label3(46)"
      Tab(2).Control(5)=   "VSCashDis_Examate"
      Tab(2).Control(6)=   "Command2"
      Tab(2).Control(7)=   "txtDepositScNote"
      Tab(2).Control(8)=   "txtExtraDiscountDet"
      Tab(2).Control(9)=   "txtExtraSPNote"
      Tab(2).Control(10)=   "txtRetPolicy"
      Tab(2).Control(11)=   "txtRetPolicyDet"
      Tab(2).Control(12)=   "txtExtraDisSc"
      Tab(2).Control(13)=   "txtExtraDisAlpha"
      Tab(2).Control(14)=   "txtRetPolicyAlpha"
      Tab(2).Control(15)=   "Check1_ExtraDis"
      Tab(2).Control(16)=   "Check1_returnp"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Page - 4"
      TabPicture(3)   =   "frmAgreement.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtbName"
      Tab(3).Control(1)=   "Command1"
      Tab(3).Control(2)=   "txtgenTerm_a"
      Tab(3).Control(3)=   "txtgenTerm_b"
      Tab(3).Control(4)=   "txtgenTerm_c"
      Tab(3).Control(5)=   "txtgenTerm_d"
      Tab(3).Control(6)=   "txtgenTerm_e"
      Tab(3).Control(7)=   "txtGeneral_Alpha"
      Tab(3).Control(8)=   "txtgenTerm_f"
      Tab(3).Control(9)=   "txtgenTerm_g"
      Tab(3).Control(10)=   "txtgenTerm_h"
      Tab(3).Control(11)=   "txtgenTerm_i"
      Tab(3).Control(12)=   "txtgenTerm_j"
      Tab(3).Control(13)=   "txtgenTerm_k"
      Tab(3).Control(14)=   "txtgenTerm_l"
      Tab(3).Control(15)=   "txtgenTerm_m"
      Tab(3).Control(16)=   "txtAboveBus"
      Tab(3).Control(17)=   "Label3(32)"
      Tab(3).Control(18)=   "Label3(33)"
      Tab(3).Control(19)=   "Label3(5)"
      Tab(3).Control(20)=   "Label3(6)"
      Tab(3).Control(21)=   "Label3(7)"
      Tab(3).Control(22)=   "Label3(34)"
      Tab(3).Control(23)=   "Label3(3)"
      Tab(3).Control(24)=   "Label3(4)"
      Tab(3).Control(25)=   "Label3(10)"
      Tab(3).Control(26)=   "Label3(36)"
      Tab(3).Control(27)=   "Label3(37)"
      Tab(3).Control(28)=   "Label3(39)"
      Tab(3).Control(29)=   "Label3(40)"
      Tab(3).Control(30)=   "Label3(35)"
      Tab(3).ControlCount=   31
      Begin VB.CommandButton cmdRepQty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agreement Litst to Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   11772
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   1656
         Width           =   2220
      End
      Begin VB.ComboBox Combo1_State 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "frmAgreement.frx":007C
         Left            =   -73920
         List            =   "frmAgreement.frx":0089
         TabIndex        =   141
         Top             =   5004
         Width           =   840
      End
      Begin VB.TextBox txtCashDepRem1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   324
         Left            =   -72696
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   140
         Top             =   5904
         Width           =   14292
      End
      Begin VB.TextBox txtTurnOverSp 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   510
         Left            =   -72696
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   139
         Top             =   2232
         Width           =   14292
      End
      Begin VB.TextBox txtTurnOverSp1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   324
         Left            =   -72696
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   138
         Top             =   2796
         Width           =   14292
      End
      Begin VB.CheckBox Check3_BaseDis 
         Caption         =   "Base Discount Term"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   408
         Left            =   12060
         TabIndex        =   137
         Top             =   3528
         Width           =   2412
      End
      Begin VB.CheckBox Check2_CD 
         Caption         =   "CD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   408
         Left            =   10800
         TabIndex        =   136
         Top             =   3528
         Width           =   1116
      End
      Begin VB.CheckBox Check1_TOD 
         Caption         =   "ASB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   408
         Left            =   9324
         TabIndex        =   135
         Top             =   3492
         Width           =   1188
      End
      Begin VB.Frame Frame1 
         Height          =   552
         Left            =   5292
         TabIndex        =   134
         Top             =   648
         Width           =   3144
         Begin VB.OptionButton Option1_new 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   2
            Top             =   180
            Width           =   984
         End
         Begin VB.OptionButton Option2_Existing 
            Caption         =   "Existing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   1584
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   1164
         End
      End
      Begin VB.CommandButton Commandedit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   645
         Left            =   11808
         Picture         =   "frmAgreement.frx":009A
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   960
         Width           =   1056
      End
      Begin VB.CheckBox Check1_OldNew 
         Caption         =   "Old Content"
         Height          =   555
         Left            =   15336
         TabIndex        =   132
         Top             =   4212
         Visible         =   0   'False
         Width           =   1404
      End
      Begin VB.TextBox txtbName 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   372
         Left            =   -74712
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   131
         Top             =   9144
         Width           =   16356
      End
      Begin VB.TextBox txtTransportation1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   480
         Left            =   -72705
         MaxLength       =   350
         MultiLine       =   -1  'True
         TabIndex        =   130
         Top             =   3852
         Width           =   14328
      End
      Begin VB.TextBox txtTurnOverSp3 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   324
         Left            =   -72696
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   129
         Top             =   1764
         Width           =   14292
      End
      Begin VB.CheckBox Check1_returnp 
         Caption         =   "Return Policy"
         Height          =   555
         Left            =   -74730
         TabIndex        =   128
         Top             =   4644
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.CheckBox Check1_ExtraDis 
         Caption         =   "Extra Dis."
         Height          =   555
         Left            =   -74730
         TabIndex        =   127
         Top             =   2592
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CheckBox Check1_TDExamte 
         Caption         =   "Turnover Discount (Examate)"
         Height          =   555
         Left            =   -73965
         TabIndex        =   126
         Top             =   8448
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txtCashDepositAlpha 
         BackColor       =   &H00C0FFC0&
         Height          =   330
         Left            =   -74820
         TabIndex        =   125
         Text            =   "E."
         Top             =   7812
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtRetPolicyAlpha 
         BackColor       =   &H00C0FFC0&
         Height          =   330
         Left            =   -74865
         TabIndex        =   124
         Text            =   "E."
         Top             =   4044
         Width           =   315
      End
      Begin VB.TextBox txtExtraDisAlpha 
         BackColor       =   &H00C0FFC0&
         Height          =   330
         Left            =   -74865
         TabIndex        =   123
         Text            =   "F."
         Top             =   2112
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtExtraDisSc 
         Height          =   375
         Left            =   -73875
         TabIndex        =   122
         Top             =   2112
         Visible         =   0   'False
         Width           =   12555
      End
      Begin VB.TextBox txtRetPolicyDet 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   1035
         Left            =   -73875
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   120
         Top             =   5856
         Width           =   13824
      End
      Begin VB.TextBox txtRetPolicy 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   990
         Left            =   -73875
         MaxLength       =   600
         MultiLine       =   -1  'True
         TabIndex        =   118
         Top             =   4356
         Width           =   13800
      End
      Begin VB.TextBox txtExtraSPNote 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   324
         Left            =   -73848
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   116
         Top             =   2808
         Visible         =   0   'False
         Width           =   12570
      End
      Begin VB.TextBox txtExtraDiscountDet 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   288
         Left            =   -73875
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   115
         Top             =   2508
         Visible         =   0   'False
         Width           =   12570
      End
      Begin VB.TextBox txtDepositScNote 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   288
         Left            =   -73875
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   113
         Top             =   1764
         Visible         =   0   'False
         Width           =   12570
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   615
         Left            =   -62688
         Picture         =   "frmAgreement.frx":04DC
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   3312
         Width           =   1320
      End
      Begin VB.TextBox txtCashDepositSch_Heading 
         Height          =   375
         Left            =   -74460
         TabIndex        =   109
         Top             =   7812
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   660
         Left            =   -62220
         Picture         =   "frmAgreement.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   360
         Width           =   1644
      End
      Begin VB.TextBox txtgenTerm_a 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   480
         Left            =   -74685
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   100
         Top             =   1050
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_b 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         DataField       =   "Name"
         Height          =   480
         Left            =   -74685
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   99
         Top             =   1530
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_c 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   480
         Left            =   -74685
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   98
         Top             =   2010
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_d 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         DataField       =   "Name"
         Height          =   480
         Left            =   -74685
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   97
         Top             =   2490
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_e 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   480
         Left            =   -74685
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   96
         Top             =   2970
         Width           =   16356
      End
      Begin VB.TextBox txtGeneral_Alpha 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   -74865
         TabIndex        =   95
         Text            =   "F."
         Top             =   540
         Width           =   300
      End
      Begin VB.TextBox txtgenTerm_f 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   555
         Left            =   -74685
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   86
         Top             =   3510
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_g 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   600
         Left            =   -74685
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   85
         Top             =   4125
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_h 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   660
         Left            =   -74700
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   4785
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_i 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   600
         Left            =   -74685
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   83
         Top             =   5505
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_j 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   960
         Left            =   -74700
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   82
         Top             =   6165
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_k 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   660
         Left            =   -74700
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   81
         Top             =   7185
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_l 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   600
         Left            =   -74685
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   80
         Top             =   7875
         Width           =   16356
      End
      Begin VB.TextBox txtgenTerm_m 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   504
         Left            =   -74700
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   79
         Top             =   8565
         Width           =   16356
      End
      Begin VB.TextBox txtAboveBus 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   336
         Left            =   -74700
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   9576
         Width           =   16356
      End
      Begin VB.TextBox txtTrans_Alfabat 
         BackColor       =   &H00C0FFC0&
         Height          =   330
         Left            =   -74820
         TabIndex        =   77
         Text            =   "C."
         Top             =   3444
         Width           =   315
      End
      Begin VB.TextBox txtSPNoteForExamation 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   288
         Left            =   2256
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   9876
         Visible         =   0   'False
         Width           =   14184
      End
      Begin VB.TextBox txtTransportation 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   444
         Left            =   -72705
         MaxLength       =   450
         MultiLine       =   -1  'True
         TabIndex        =   72
         Top             =   3372
         Width           =   14328
      End
      Begin VB.CommandButton Commanddelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   645
         Left            =   12888
         Picture         =   "frmAgreement.frx":1CA4
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   960
         Width           =   1056
      End
      Begin VB.CommandButton CommandReturn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Return"
         Height          =   645
         Left            =   15024
         Picture         =   "frmAgreement.frx":2888
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   960
         Width           =   1056
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   75
         Left            =   -74700
         TabIndex        =   69
         Top             =   4320
         Width           =   75
         _ExtentX        =   127
         _ExtentY        =   127
         _Version        =   393216
      End
      Begin VB.TextBox txtCashDis_Alpha 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   -74820
         TabIndex        =   68
         Text            =   "D."
         Top             =   4344
         Width           =   315
      End
      Begin VB.TextBox txtTurnDis_Alpha 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   -74820
         TabIndex        =   67
         Text            =   "B."
         Top             =   465
         Width           =   315
      End
      Begin VB.TextBox txtCashMinExtra1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   66
         Top             =   6984
         Width           =   1635
      End
      Begin VB.TextBox txtCashMinExtra2 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   65
         Top             =   7344
         Width           =   1635
      End
      Begin VB.TextBox txtCashMinAmt2 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   -71520
         MaxLength       =   20
         TabIndex        =   59
         Top             =   7344
         Width           =   1635
      End
      Begin VB.TextBox txtCashMinAmt1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   -71520
         MaxLength       =   20
         TabIndex        =   57
         Top             =   6984
         Width           =   1635
      End
      Begin VB.CheckBox Check1_cashdep 
         Caption         =   "Cash Deposit"
         Height          =   375
         Left            =   -73920
         TabIndex        =   56
         Top             =   5400
         Width           =   990
      End
      Begin VB.TextBox txtCashDepSpNot 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   324
         Left            =   -72705
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   54
         Top             =   6480
         Width           =   14292
      End
      Begin VB.TextBox txtCashDepoSch 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   312
         Left            =   -72705
         MaxLength       =   400
         TabIndex        =   51
         Top             =   4392
         Width           =   14328
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   660
         Left            =   -59364
         Picture         =   "frmAgreement.frx":346C
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   900
         Width           =   960
      End
      Begin VB.CheckBox Check1_turnOver 
         Caption         =   "ASB"
         Height          =   555
         Left            =   -73875
         TabIndex        =   49
         Top             =   1296
         Width           =   975
      End
      Begin VB.TextBox txtTurnOverDis 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   336
         Left            =   -72705
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   465
         Width           =   13176
      End
      Begin Crystal.CrystalReport cr 
         Left            =   450
         Top             =   7605
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton CommandPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   645
         Left            =   13956
         Picture         =   "frmAgreement.frx":4050
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   960
         Width           =   1056
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   645
         Left            =   10716
         Picture         =   "frmAgreement.frx":4C34
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   960
         Width           =   1056
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   645
         Left            =   9648
         Picture         =   "frmAgreement.frx":5818
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   960
         Width           =   1056
      End
      Begin VB.TextBox txtSession1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   312
         Left            =   11070
         MaxLength       =   100
         TabIndex        =   40
         Top             =   4860
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtSpNoteC 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   336
         Left            =   2256
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   9516
         Visible         =   0   'False
         Width           =   14184
      End
      Begin VB.TextBox txtSpNoteB 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   285
         Left            =   2256
         MaxLength       =   200
         TabIndex        =   36
         Top             =   9204
         Width           =   14184
      End
      Begin VB.TextBox txtSpNoteA 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   555
         Left            =   2256
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   8616
         Width           =   14184
      End
      Begin VB.TextBox txtexpSale 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   2250
         MaxLength       =   100
         TabIndex        =   32
         Top             =   5670
         Width           =   2130
      End
      Begin VB.TextBox txtarea 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   2250
         MaxLength       =   100
         TabIndex        =   30
         Top             =   5310
         Width           =   2130
      End
      Begin VB.TextBox txtwarm3 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   348
         Left            =   1935
         MaxLength       =   150
         TabIndex        =   29
         Text            =   "As per our mutual discussion, we are pleased to share the following business terms for the session "
         Top             =   4860
         Width           =   8868
      End
      Begin VB.TextBox txtwarm2 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   540
         Left            =   1908
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "frmAgreement.frx":63FC
         Top             =   4230
         Width           =   13380
      End
      Begin VB.TextBox txtwarm1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   285
         Left            =   1935
         MaxLength       =   100
         TabIndex        =   27
         Text            =   "Warm Greeting from Blueprint Education !"
         Top             =   3915
         Width           =   6468
      End
      Begin VB.TextBox txtdearsir 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   285
         Left            =   1935
         MaxLength       =   100
         TabIndex        =   26
         Text            =   "Dear Sir,"
         Top             =   3600
         Width           =   2265
      End
      Begin VB.TextBox txtSub_yrs 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   4860
         MaxLength       =   100
         TabIndex        =   25
         Top             =   3240
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   1935
         MaxLength       =   100
         TabIndex        =   23
         Text            =   "MOU for the year"
         Top             =   3240
         Width           =   2265
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   9312
         MaxLength       =   100
         TabIndex        =   22
         Top             =   2835
         Width           =   3120
      End
      Begin VB.TextBox txtMobile 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   9312
         MaxLength       =   100
         TabIndex        =   20
         Top             =   2430
         Width           =   3120
      End
      Begin VB.TextBox txtAddress4 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   1935
         MaxLength       =   100
         TabIndex        =   18
         Top             =   2880
         Width           =   6468
      End
      Begin VB.TextBox txtAddress3 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   1935
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2520
         Width           =   6468
      End
      Begin VB.TextBox txtAddress2 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   1935
         MaxLength       =   100
         TabIndex        =   14
         Top             =   2160
         Width           =   6468
      End
      Begin VB.TextBox txtAddress1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   1935
         MaxLength       =   150
         TabIndex        =   12
         Top             =   1800
         Width           =   6468
      End
      Begin VB.TextBox txtPName 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   1944
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1440
         Width           =   6468
      End
      Begin VB.TextBox txtAgmNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Pub_code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1935
         TabIndex        =   0
         Top             =   720
         Width           =   948
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   1935
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   315
         Left            =   3735
         TabIndex        =   1
         Top             =   720
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   550
         _Version        =   393216
         CalendarBackColor=   16776960
         Format          =   172163073
         CurrentDate     =   38372
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   3168
         Left            =   4500
         TabIndex        =   43
         Top             =   5424
         Width           =   11988
         _cx             =   21145
         _cy             =   5588
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   11162880
         BackColorFixed  =   7917545
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   -2147483647
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   310
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAgreement.frx":653C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid vs_page2 
         Height          =   852
         Left            =   -72708
         TabIndex        =   46
         Top             =   840
         Width           =   13176
         _cx             =   23241
         _cy             =   1503
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   11162880
         BackColorFixed  =   7917545
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   -2147483647
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAgreement.frx":662A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid vsCashDeposit 
         Height          =   1068
         Left            =   -72708
         TabIndex        =   53
         Top             =   4776
         Width           =   14328
         _cx             =   25273
         _cy             =   1884
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   11162880
         BackColorFixed  =   7917545
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   -2147483647
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAgreement.frx":6718
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid VSExamtion 
         Height          =   108
         Left            =   2256
         TabIndex        =   76
         Top             =   9852
         Visible         =   0   'False
         Width           =   10380
         _cx             =   18309
         _cy             =   190
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   11162880
         BackColorFixed  =   7917545
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   -2147483647
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAgreement.frx":6806
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin RichTextLib.RichTextBox txtCashDepositSch 
         Height          =   936
         Left            =   -72708
         TabIndex        =   108
         Top             =   8316
         Visible         =   0   'False
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   1651
         _Version        =   393217
         ScrollBars      =   3
         RightMargin     =   20000
         TextRTF         =   $"frmAgreement.frx":6858
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VSFlex7Ctl.VSFlexGrid VSCashDis_Examate 
         Height          =   576
         Left            =   -73872
         TabIndex        =   110
         Top             =   996
         Visible         =   0   'False
         Width           =   12588
         _cx             =   22204
         _cy             =   1016
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   11162880
         BackColorFixed  =   7917545
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   -2147483647
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAgreement.frx":68D8
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid vs1 
         Height          =   288
         Left            =   14004
         TabIndex        =   142
         Top             =   1944
         Visible         =   0   'False
         Width           =   1944
         _cx             =   3429
         _cy             =   508
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   11162880
         BackColorFixed  =   7917545
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   -2147483647
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   310
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAgreement.frx":695E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Return Policy Detail :"
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   46
         Left            =   -73920
         TabIndex        =   121
         Top             =   5580
         Width           =   1992
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Return Policy :"
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   45
         Left            =   -73920
         TabIndex        =   119
         Top             =   4092
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Special Note :"
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   44
         Left            =   -74712
         TabIndex        =   117
         Top             =   3060
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Special Note :"
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   42
         Left            =   -73872
         TabIndex        =   114
         Top             =   1536
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Deposit Scheme (Exam Mate):"
         Height          =   375
         Left            =   -73875
         TabIndex        =   111
         Top             =   675
         Visible         =   0   'False
         Width           =   3930
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "a."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   32
         Left            =   -74910
         TabIndex        =   106
         Top             =   1035
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "b."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   33
         Left            =   -74910
         TabIndex        =   105
         Top             =   1590
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "c."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   -74910
         TabIndex        =   104
         Top             =   2010
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "d."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   -74910
         TabIndex        =   103
         Top             =   2550
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "e."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   -74910
         TabIndex        =   102
         Top             =   2970
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "G.General Terms of Business :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   34
         Left            =   -74460
         TabIndex        =   101
         Top             =   540
         Width           =   4035
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "f."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   -74910
         TabIndex        =   94
         Top             =   3465
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "g."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   -74910
         TabIndex        =   93
         Top             =   4125
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "h."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   -74910
         TabIndex        =   92
         Top             =   4785
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "i."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   36
         Left            =   -74910
         TabIndex        =   91
         Top             =   5505
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "j."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   37
         Left            =   -74910
         TabIndex        =   90
         Top             =   6165
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "k."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   39
         Left            =   -74910
         TabIndex        =   89
         Top             =   7185
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "l."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   40
         Left            =   -74910
         TabIndex        =   88
         Top             =   7905
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "m."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   35
         Left            =   -74910
         TabIndex        =   87
         Top             =   8565
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Special Note :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   41
         Left            =   900
         TabIndex        =   75
         Top             =   9780
         Visible         =   0   'False
         Width           =   1368
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Transportation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   288
         Index           =   38
         Left            =   -74508
         TabIndex        =   73
         Top             =   3480
         Width           =   1632
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Extra % other than Scheme-2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   276
         Index           =   31
         Left            =   -69840
         TabIndex        =   64
         Top             =   7404
         Width           =   2508
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Extra % other than Scheme-1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   276
         Index           =   30
         Left            =   -69840
         TabIndex        =   63
         Top             =   7044
         Width           =   2508
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimun Amt2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   276
         Index           =   27
         Left            =   -72720
         TabIndex        =   60
         Top             =   7344
         Width           =   1428
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimun Amt1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   276
         Index           =   26
         Left            =   -72720
         TabIndex        =   58
         Top             =   6984
         Width           =   1428
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Special Note :"
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   25
         Left            =   -73740
         TabIndex        =   55
         Top             =   6504
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "F. Cash Deposit Scheme :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   528
         Index           =   9
         Left            =   -74448
         TabIndex        =   52
         Top             =   4368
         Width           =   1920
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Special Note :"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   -73800
         TabIndex        =   48
         Top             =   2370
         Width           =   1080
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Settlement Bonus (ASB) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   432
         Index           =   0
         Left            =   -74472
         TabIndex        =   47
         Top             =   468
         Width           =   1872
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "D. Discount Structure - Examination Material"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   132
         Index           =   24
         Left            =   360
         TabIndex        =   39
         Top             =   9792
         Visible         =   0   'False
         Width           =   1776
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Special Note :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   23
         Left            =   2256
         TabIndex        =   38
         Top             =   8208
         Width           =   1680
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "C. Discount Structure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   288
         Index           =   22
         Left            =   4500
         TabIndex        =   35
         Top             =   5208
         Width           =   1908
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "B. Expected Sales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   21
         Left            =   360
         TabIndex        =   33
         Top             =   5715
         Width           =   1680
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A. Area of Operation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   20
         Left            =   360
         TabIndex        =   31
         Top             =   5355
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   19
         Left            =   360
         TabIndex        =   24
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   18
         Left            =   8556
         TabIndex        =   21
         Top             =   2928
         Width           =   1140
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   288
         Index           =   17
         Left            =   8556
         TabIndex        =   19
         Top             =   2472
         Width           =   1056
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   16
         Left            =   360
         TabIndex        =   17
         Top             =   2925
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   15
         Left            =   360
         TabIndex        =   15
         Top             =   2565
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   " Address2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   14
         Left            =   315
         TabIndex        =   13
         Top             =   2205
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   13
         Left            =   360
         TabIndex        =   11
         Top             =   1845
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Agreement No  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   12
         Left            =   405
         TabIndex        =   10
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   8
         Left            =   2970
         TabIndex        =   9
         Top             =   765
         Width           =   705
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   " Party Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   315
         TabIndex        =   8
         Top             =   1485
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   1125
         Width           =   1425
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimun Amt2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   29
      Left            =   5220
      TabIndex        =   62
      Top             =   7320
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimun Amt1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   28
      Left            =   5220
      TabIndex        =   61
      Top             =   6645
      Width           =   1425
   End
End
Attribute VB_Name = "frmAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean
Private Sub Check1_cashdep_Click()
   chechBoxSet
End Sub
Sub chechBoxSet()

' Dim arr(7) As String
' arr(0) = "E."
' arr(1) = "F."
' arr(2) = "G."
' arr(3) = "H."
' arr(4) = "I."
' arr(5) = "J."
' arr(6) = "K."
'
'
'
'Dim k1 As Integer
'k1 = 0
'
'If Check1_turnOver.value = 0 Then
'   txtTurnDis_Alpha.Text = ""
'   txtTrans_Alfabat.Text = arr(k1)
'   k1 = k1 + 1
'Else
'   txtTurnDis_Alpha.Text = arr(k1)
'   k1 = k1 + 1
'   txtTrans_Alfabat.Text = arr(k1)
'   k1 = k1 + 1
'
'End If
'
'
'If Check1_cashdep.value = 0 Then
'   txtCashDis_Alpha.Text = ""
'Else
'   txtCashDis_Alpha.Text = arr(k1)
'   k1 = k1 + 1
'End If
'
'If Check1_TDExamte.value = 0 Then
'   txtCashDepositAlpha.Text = ""
'Else
'   txtCashDepositAlpha.Text = arr(k1)
'   k1 = k1 + 1
'End If
'
'If Check1_ExtraDis.value = 0 Then
'   txtExtraDisAlpha.Text = ""
'Else
'   txtExtraDisAlpha.Text = arr(k1)
'   k1 = k1 + 1
'End If
'
'
'If Check1_returnp.value = 0 Then
'   txtRetPolicyAlpha.Text = ""
'Else
'   txtRetPolicyAlpha.Text = arr(k1)
'   k1 = k1 + 1
'End If
'
'
'
'txtGeneral_Alpha.Text = arr(k1)
'k1 = k1 + 1
'

 
 
End Sub

Private Sub Check1_ExtraDis_Click()
chechBoxSet
End Sub

Private Sub Check1_OldNew_Click()

If Check1_OldNew.value = 1 Then

txtwarm2.text = "We are pleased to have received your customer agreement form and appreciate the interest that you have shown in our books. We thank you for the courtesy and cooperation extended at your end. We strongly believe that you are one of the best entrepreneurs of your region and our business association will be beneficial for schools that are looking for quality books for their students."

Else

txtwarm2.text = "Thank you for your support last year.  We value our partnership and believe your leadership will continue benefiting schools in need of quality books."

End If

End Sub

Private Sub Check1_returnp_Click()
chechBoxSet
End Sub

Private Sub Check1_TDExamte_Click()
chechBoxSet
End Sub

Private Sub Check1_TOD_Click()



If Check1_TOD.value = 1 Then
   Check1_turnOver.value = 1
Else
   Check1_turnOver.value = 0
End If

View_Hide

End Sub
Sub View_Hide()


txtTurnDis_Alpha.text = "B."
txtTrans_Alfabat.text = "C."
txtCashDis_Alpha.text = "D."
txtRetPolicyAlpha.text = "E."


If Check1_TOD.value = 1 And Check2_CD.value = 1 Then

txtCashDis_Alpha.text = "D."
txtRetPolicyAlpha.text = "E."

txtGeneral_Alpha.text = "F."


ElseIf Check1_TOD.value = 1 And Check2_CD.value = 0 Then

txtCashDis_Alpha.text = "D."
txtRetPolicyAlpha.text = "D."
txtGeneral_Alpha.text = "E."

ElseIf Check1_TOD.value = 0 And Check2_CD.value = 1 Then

txtTrans_Alfabat.text = "B."
txtCashDis_Alpha.text = "C."
txtRetPolicyAlpha.text = "D."
txtGeneral_Alpha.text = "E."

ElseIf Check1_TOD.value = 0 And Check2_CD.value = 0 Then

txtTrans_Alfabat.text = "B."

txtRetPolicyAlpha.text = "C."

txtGeneral_Alpha.text = "D."

End If





End Sub
Private Sub Check1_turnOver_Click()
   chechBoxSet
End Sub

Private Sub Check2_CD_Click()

If Check2_CD.value = 1 Then
   Check1_cashdep.value = 1
Else
   Check1_cashdep.value = 0
End If

View_Hide

End Sub

Private Sub cmdAdd_1_Click()

Edit = False
txtAgmNo = MaxSNo("AgreementMain", "agmno")

Option2_Existing.value = True

iniForm

txtName.text = ""
txtPName.text = ""

txtAddress1.text = ""
txtAddress2.text = ""
txtAddress3.text = ""
txtAddress4.text = ""

txtMobile.text = ""
txtEmail = ""

txtArea.text = ""
txtexpSale.text = ""

Check3_BaseDis.value = 0

Commanddelete.Enabled = False
Commandedit.Enabled = False
cmdSave_2.Enabled = True



End Sub
Sub searchData()

On Error GoTo search_

If RS.State = 1 Then RS.close
RS.Open "select * from AgreementMain where AgmNo=" & txtAgmNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
   Exit Sub
End If

If RS.EOF = False Then
  

 Commanddelete.Enabled = False
 Commandedit.Enabled = True
 cmdSave_2.Enabled = False
  
 Combo1_State.text = RS!examdesc2 & ""
  
  
 If RS!New_Existing = "n" Then
    Option1_new.value = True
 Else
   Option2_Existing.value = True
 End If
  
  
 txtAgmNo.text = RS!AgmNo
 txtDate.value = RS!dates
 txtName.text = RS!Name
 txtPName.text = RS!pname

 txtAddress1.text = RS!address1
 txtAddress2.text = RS!address2
 txtAddress3.text = RS!address3
 txtAddress4.text = RS!address4

 txtMobile.text = RS!mobile
 txtEmail = RS!email

 txtSubject.text = RS!Subject1
 txtSub_yrs.text = RS!Subject2

 txtdearsir.text = RS!DearSir

 txtwarm1 = RS!WarmGreeting
 txtwarm2 = RS!WarmGreetingText1
 txtwarm3 = RS!WarmGreetingText2

 txtArea = RS!Area
 txtexpSale = RS!ExpSale

 txtSpNoteA = RS!SpNote_a
 txtSpNoteB = RS!SpNote_b
 txtSpNoteC = RS!SpNote_c

 txtTransportation = RS!Transportation & ""
 
 
 VSExamtion.TextMatrix(1, 0) = RS!examdesc1 & ""
 VSExamtion.TextMatrix(2, 0) = RS!examdesc2 & ""

 VSExamtion.TextMatrix(1, 1) = RS!Examclass1 & ""
 VSExamtion.TextMatrix(2, 1) = RS!Examclass2 & ""

 VSExamtion.TextMatrix(1, 2) = RS!ExamPercentage1 & ""
 VSExamtion.TextMatrix(2, 2) = RS!ExamPercentage2 & ""
 txtSPNoteForExamation.text = RS!ExamSpNote & ""
 
 txtTrans_Alfabat.text = RS!TransportAlfa & ""
 
 '===========================Page -2 ===============================
 If RS!turnOver_yesno = "n" Then
    Check1_turnOver.value = 0
 Else
    Check1_turnOver.value = 1
 End If

 txtTurnOverDis.text = RS!turnOverDis & ""
 
 vs_page2.TextMatrix(1, 0) = RS!turnOverPart1 & ""
' vs_page2.TextMatrix(1, 1) = RS!turnOverDis1 & ""
' vs_page2.TextMatrix(2, 0) = RS!turnOverPart2 & ""
' vs_page2.TextMatrix(2, 1) = RS!turnOverDis2 & ""
' vs_page2.TextMatrix(3, 0) = RS!turnOverPart3 & ""
' vs_page2.TextMatrix(3, 1) = RS!turnOverDis3 & ""
' vs_page2.TextMatrix(4, 0) = RS!turnOverPart4 & ""
' vs_page2.TextMatrix(4, 1) = RS!turnOverDis4 & ""
' vs_page2.TextMatrix(5, 0) = RS!turnOverPart5 & ""
' vs_page2.TextMatrix(5, 1) = RS!turnOverDis5 & ""
' vs_page2.TextMatrix(6, 0) = RS!turnOverPart6 & ""
' vs_page2.TextMatrix(6, 1) = RS!turnOverDis6 & ""
 
    vs_page2.TextMatrix(1, 0) = RS!turnOverPart1 & ""
    vs_page2.TextMatrix(1, 1) = RS!turnOverPart2 & ""
    vs_page2.TextMatrix(1, 2) = RS!turnOverPart3 & ""
    vs_page2.TextMatrix(1, 3) = RS!turnOverPart4 & ""
    vs_page2.TextMatrix(1, 4) = RS!turnOverPart5 & ""

 

 txtTurnOverSp.text = RS!turnOver_SpNote & ""
 txtTurnOverSp1.text = RS!turnOver_SpNote1 & ""

 
'===========================End Page-2=============================
 
    
txtCashDepoSch.text = RS!CashDepositSh & ""
    
vsCashDeposit.TextMatrix(0, 1) = RS!CashExamateSep1 & ""
vsCashDeposit.TextMatrix(0, 2) = RS!CashExamateSep2 & ""
vsCashDeposit.TextMatrix(0, 3) = RS!CashExamateSep3 & ""
vsCashDeposit.TextMatrix(0, 4) = RS!CashExamateSep4 & ""
    
    
vsCashDeposit.TextMatrix(1, 0) = RS!CashDepositPeriod1 & ""
vsCashDeposit.TextMatrix(1, 1) = RS!CashDepositExtraCr1 & ""
vsCashDeposit.TextMatrix(1, 2) = RS!CashDepositPeriod2 & ""
vsCashDeposit.TextMatrix(1, 3) = RS!CashDepositExtraCr2 & ""
vsCashDeposit.TextMatrix(1, 4) = RS!CashDepositPeriod3 & ""
vsCashDeposit.TextMatrix(2, 4) = RS!CashDepositPeriod4 & ""
    
    txtCashDepSpNot.text = RS!CashDepositSpNot & ""
    txtCashDepRem1.text = RS!examdesc1 & ""
    
    txtCashMinAmt1.text = RS!CashDepositMinAmt1 & ""
    txtCashMinAmt2.text = RS!CashDepositMinAmt2 & ""
    
    txtCashMinExtra1.text = RS!CashDepositMinExtra1 & ""
    txtCashMinExtra2.text = RS!CashDepositMinExtra2 & ""
 
  
 '==================================================================
 
 txtgenTerm_a.text = RS!GenTerms_a & ""
 txtgenTerm_b.text = RS!GenTerms_b & ""
 txtgenTerm_c.text = RS!GenTerms_c & ""
 txtgenTerm_d.text = RS!GenTerms_d & ""
 txtgenTerm_e.text = RS!GenTerms_e & ""
 
 '==================================================================
 
 txtgenTerm_f.text = RS!GenTerms_f & ""
 txtgenTerm_g.text = RS!GenTerms_g & ""
 txtgenTerm_h.text = RS!GenTerms_h & ""
 txtgenTerm_i.text = RS!GenTerms_i & ""
 txtgenTerm_j.text = RS!GenTerms_j & ""
 txtgenTerm_k.text = RS!GenTerms_k & ""
 txtgenTerm_l.text = RS!GenTerms_l & ""
 txtgenTerm_m.text = RS!GenTerms_m & ""
 txtAboveBus.text = RS!AboveTerms & ""
 
 txtTurnDis_Alpha.text = RS!turnOverINI & ""
 txtCashDis_Alpha.text = RS!CashDepositINI & ""
 txtGeneral_Alpha.text = RS!GenTermINI & ""
 
 txtTrans_Alfabat.text = RS!TransportAlfa & ""
 
 
 txtCashDepositSch.text = RS!CashDepositSc & ""
 txtCashDepositSch_Heading = RS!CashDepositSch_Heading & ""
 
 
 
 
VSCashDis_Examate.TextMatrix(1, 0) = RS!CashExamateAmt1 & ""
VSCashDis_Examate.TextMatrix(2, 0) = RS!CashExamateAmt2 & ""
VSCashDis_Examate.TextMatrix(3, 0) = RS!CashExamateAmt3 & ""
VSCashDis_Examate.TextMatrix(4, 0) = RS!CashExamateAmt4 & ""
VSCashDis_Examate.TextMatrix(5, 0) = RS!CashExamateAmt5 & ""

VSCashDis_Examate.TextMatrix(1, 1) = RS!CashExamateMay1 & ""
VSCashDis_Examate.TextMatrix(2, 1) = RS!CashExamateMay2 & ""
VSCashDis_Examate.TextMatrix(3, 1) = RS!CashExamateMay3 & ""
VSCashDis_Examate.TextMatrix(4, 1) = RS!CashExamateMay4 & ""
VSCashDis_Examate.TextMatrix(5, 1) = RS!CashExamateMay5 & ""

VSCashDis_Examate.TextMatrix(1, 2) = RS!CashExamateJul1 & ""
VSCashDis_Examate.TextMatrix(2, 2) = RS!CashExamateJul2 & ""
VSCashDis_Examate.TextMatrix(3, 2) = RS!CashExamateJul3 & ""
VSCashDis_Examate.TextMatrix(4, 2) = RS!CashExamateJul4 & ""
VSCashDis_Examate.TextMatrix(5, 2) = RS!CashExamateJul5 & ""

VSCashDis_Examate.TextMatrix(1, 3) = RS!CashExamateSep1 & ""
VSCashDis_Examate.TextMatrix(2, 3) = RS!CashExamateSep2 & ""
VSCashDis_Examate.TextMatrix(3, 3) = RS!CashExamateSep3 & ""
VSCashDis_Examate.TextMatrix(4, 3) = RS!CashExamateSep4 & ""
VSCashDis_Examate.TextMatrix(5, 3) = RS!CashExamatesep5 & ""

txtDepositScNote.text = RS!DepositSpNot & ""

txtExtraDisAlpha.text = RS!ExtraDisAlpha & ""
'txtExtraDisSc.text = RS!extraDisSc & ""
txtExtraDiscountDet.text = RS!ExtraDisScDet & ""
txtExtraSPNote.text = RS!ExtraDisScNot & ""


txtRetPolicy.text = RS!RetPolicy & ""
txtRetPolicyDet.text = RS!RetPolicyDet & ""

txtCashDepositAlpha.text = RS!CashDepositAlpha & ""
txtExtraDisAlpha.text = RS!ExtraDisAlpha & ""

txtRetPolicyAlpha.text = RS!RetPolicyAlpha & ""

txtbName.text = RS!bankAcc & ""


If RS!TDExamate_y_n = "y" Then
   Check1_TDExamte.value = 1
Else
   Check1_TDExamte.value = 0
End If

If Len(RS!ExtraDisAlpha) = 2 Then
   Check1_ExtraDis.value = 1
Else
   Check1_ExtraDis.value = 0
End If

If Len(RS!RetPolicyAlpha) = 2 Then
   Check1_returnp.value = 1
Else
   Check1_returnp.value = 0
End If


If RS!extraDisSc = "y" Then
   Check3_BaseDis.value = 1
Else
   Check3_BaseDis.value = 0
End If


If RS!CashDepositYesNo = "n" Then
     Check1_cashdep.value = 0
     Check2_CD.value = 0
 Else
     Check1_cashdep.value = 1
     Check2_CD.value = 1
End If

If RS!turnOver_yesno = "n" Then
   Check1_TOD.value = 0
   Check1_turnOver.value = 0
Else
   Check1_TOD.value = 1
   Check1_turnOver.value = 1
End If


 
 If Check1_TOD.value = 1 Then
    Check1_turnOver.value = 1
 Else
    Check1_turnOver.value = 0
 End If
 
 If Check2_CD.value = 1 Then
    Check1_cashdep.value = 1
 Else
    Check1_cashdep.value = 0
 End If
 

End If


vs.Clear


If RS.State = 1 Then RS.close
RS.Open "select * from AgreementDisStructure where AgmNo=" & txtAgmNo & " order by sn", con, adOpenDynamic, adLockOptimistic

For I = 1 To RS.RecordCount

vs.TextMatrix(I, 0) = RS!sn
vs.TextMatrix(I, 1) = RS!DESCRIPTION
vs.TextMatrix(I, 2) = RS!Class
vs.TextMatrix(I, 3) = RS!Percent
vs.TextMatrix(I, 4) = RS!DisOffered & ""

RS.MoveNext

Next


fillGrid


Exit Sub

search_:

MsgBox "" & err.DESCRIPTION


End Sub

Private Sub cmdRepQty_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String

Screen.MousePointer = vbHourglass

If rs_1.State = 1 Then rs_1.close
rs_1.Open "SELECT [AgmNo],LTrim(substring(PName,1,5)) as Code,LTrim(substring(PName,7,100)) as Party,Address2 as City,Address3 as States,turnOver_yesno as ASB,CashDepositYesNo as CD FROM AgreementMain order by AgmNo", con

Set vs1.DataSource = rs_1



If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add


'Dim c, r As Long
'Dim Q1, q2, J, Q1_sp As Double

 
  
 
  row_ = 2
  col_ = 1
   
    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    xlSheet.Cells(1, 1).value = "AgmNo"
    xlSheet.Cells(1, 2).value = "Code"
    xlSheet.Cells(1, 3).value = "PName"
    xlSheet.Cells(1, 4).value = "City"
    xlSheet.Cells(1, 5).value = "State"
    xlSheet.Cells(1, 6).value = "ASB"
    xlSheet.Cells(1, 7).value = "CD"
    
    For I = 0 To vs1.rows - 1
        For J = 0 To vs1.Cols - 1
               xlSheet.Cells(row_, col_).value = vs1.TextMatrix(I, J)
              col_ = col_ + 1
        Next
        row_ = row_ + 1
        col_ = 1
    Next
    
    Screen.MousePointer = vbDefault
    
   

End Sub

Private Sub cmdSave_2_Click()

On Error GoTo save_



If Edit = False Then

    Set RS = New ADODB.Recordset
    RS.Open "select * from AgreementMain where AgmNo=" & txtAgmNo & "", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
   
        'MsgBox "Agreement Already Made on this Number...", vbCritical
        'Exit Sub
        txtAgmNo = MaxSNo("AgreementMain", "agmno")
    
    End If

Else


''======================================
'If Edit = False Then
'   txtAgmNo = MaxSNo("AgreementMain", "agmno")
'End If
''======================================

 Set rs1 = New ADODB.Recordset
 If rs1.State = 1 Then rs1.close
 rs1.Open "select top 1 * from AgreementMain where agmno=" & txtAgmNo.text & "", con, adOpenStatic, adLockReadOnly
 If rs1.EOF = False Then
     If rs1!bAuthorized = True Then
         MsgBox "You can'nt change, Already Locked !!", vbExclamation, "Alert"
         Exit Sub
     End If
    
 End If

'=======================================


End If




Set RS = New ADODB.Recordset
RS.Open "select * from AgreementMain where AgmNo=" & txtAgmNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
RS.AddNew
End If

RS!AgmNo = txtAgmNo.text
RS!dates = txtDate.value
RS!Name = txtName.text
RS!pname = txtPName.text

RS!address1 = txtAddress1.text
RS!address2 = txtAddress2.text
RS!address3 = txtAddress3.text
RS!address4 = txtAddress4.text

RS!mobile = txtMobile.text
RS!email = txtEmail

RS!Subject1 = txtSubject.text
RS!Subject2 = txtSub_yrs.text

RS!DearSir = txtdearsir.text

RS!WarmGreeting = txtwarm1
RS!WarmGreetingText1 = txtwarm2
RS!WarmGreetingText2 = txtwarm3

RS!Area = txtArea
RS!ExpSale = txtexpSale

RS!SpNote_a = txtSpNoteA
RS!SpNote_b = txtSpNoteB
RS!SpNote_c = txtSpNoteC

RS!examdesc1 = VSExamtion.TextMatrix(1, 0)
RS!examdesc2 = VSExamtion.TextMatrix(2, 0)

RS!Examclass1 = VSExamtion.TextMatrix(1, 1)
RS!Examclass2 = VSExamtion.TextMatrix(2, 1)

RS!ExamPercentage1 = VSExamtion.TextMatrix(1, 2)
RS!ExamPercentage2 = VSExamtion.TextMatrix(2, 2)


RS!ExamSpNote = Trim(txtSPNoteForExamation.text)

If Option1_new.value = True Then
   RS!New_Existing = "n"
Else
   RS!New_Existing = "e"
End If

If Check3_BaseDis.value = 1 Then
   RS!extraDisSc = "y"
Else
   RS!extraDisSc = "n"
End If


RS.update

'============================================================
con.Execute "delete from AgreementDisStructure where AgmNo=" & txtAgmNo & ""

Set RS = New ADODB.Recordset
RS.Open "select * from AgreementDisStructure", con, adOpenDynamic, adLockOptimistic


For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 1) <> "" Then

RS.AddNew
RS!AgmNo = txtAgmNo.text
RS!sn = vs.TextMatrix(I, 0)
RS!DESCRIPTION = vs.TextMatrix(I, 1)
RS!Class = vs.TextMatrix(I, 2)
RS!Percent = vs.TextMatrix(I, 3)
RS!DisOffered = vs.TextMatrix(I, 4)
RS.update

End If

Next


Command1_Click
Command2_Click
Command3_Click

MsgBox "Data Saved...", vbInformation




Exit Sub

save_:

MsgBox "" & err.DESCRIPTION





End Sub

Private Sub Combo1_State_Change()
fetchStateWiseCD
End Sub

Private Sub Combo1_State_Click()
fetchStateWiseCD
End Sub

Private Sub Command1_Click()

On Error GoTo save_

Set RS = New ADODB.Recordset
RS.Open "select * from AgreementMain where AgmNo=" & txtAgmNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
RS.AddNew
End If

RS!GenTerms_a = Trim(txtgenTerm_a.text)
RS!GenTerms_b = Trim(txtgenTerm_b.text)
RS!GenTerms_c = Trim(txtgenTerm_c.text)
RS!GenTerms_d = Trim(txtgenTerm_d.text)
RS!GenTerms_e = Trim(txtgenTerm_e.text)

RS!GenTerms_f = txtgenTerm_f.text
RS!GenTerms_g = txtgenTerm_g.text
RS!GenTerms_h = txtgenTerm_h.text
RS!GenTerms_i = txtgenTerm_i.text
RS!GenTerms_j = txtgenTerm_j.text
RS!GenTerms_k = txtgenTerm_k.text
RS!GenTerms_l = txtgenTerm_l.text
RS!GenTerms_m = txtgenTerm_m.text
RS!AboveTerms = txtAboveBus.text

'RS!GenTermINI = Trim(txtGeneral_Alpha.Text)
'---------------End Gen Terms==================================
RS!GenTermINI = UCase(txtGeneral_Alpha.text)

RS!bankAcc = txtbName.text






RS.update


'cmdSave_2_Click
'Command2_Click
'Command3_Click

'MsgBox "Data Saved...", vbInformation

Exit Sub


save_:

MsgBox "" & err.DESCRIPTION

End Sub



Private Sub Command2_Click()
On Error GoTo save_

Set RS = New ADODB.Recordset
RS.Open "select * from AgreementMain where AgmNo=" & txtAgmNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
RS.AddNew
End If


RS!CashExamateAmt1 = VSCashDis_Examate.TextMatrix(1, 0)
RS!CashExamateAmt2 = VSCashDis_Examate.TextMatrix(2, 0)
RS!CashExamateAmt3 = VSCashDis_Examate.TextMatrix(3, 0)
RS!CashExamateAmt4 = VSCashDis_Examate.TextMatrix(4, 0)
RS!CashExamateAmt5 = VSCashDis_Examate.TextMatrix(5, 0)

RS!CashExamateMay1 = VSCashDis_Examate.TextMatrix(1, 1)
RS!CashExamateMay2 = VSCashDis_Examate.TextMatrix(2, 1)
RS!CashExamateMay3 = VSCashDis_Examate.TextMatrix(3, 1)
RS!CashExamateMay4 = VSCashDis_Examate.TextMatrix(4, 1)
RS!CashExamateMay5 = VSCashDis_Examate.TextMatrix(5, 1)

RS!CashExamateJul1 = VSCashDis_Examate.TextMatrix(1, 2)
RS!CashExamateJul2 = VSCashDis_Examate.TextMatrix(2, 2)
RS!CashExamateJul3 = VSCashDis_Examate.TextMatrix(3, 2)
RS!CashExamateJul4 = VSCashDis_Examate.TextMatrix(4, 2)
RS!CashExamateJul5 = VSCashDis_Examate.TextMatrix(5, 2)

RS!CashExamateSep1 = VSCashDis_Examate.TextMatrix(1, 3)
RS!CashExamateSep2 = VSCashDis_Examate.TextMatrix(2, 3)
RS!CashExamateSep3 = VSCashDis_Examate.TextMatrix(3, 3)
RS!CashExamateSep4 = VSCashDis_Examate.TextMatrix(4, 3)
RS!CashExamatesep5 = VSCashDis_Examate.TextMatrix(5, 3)

RS!DepositSpNot = Trim(txtDepositScNote.text)

RS!ExtraDisAlpha = Trim(txtExtraDisAlpha.text)
'RS!extraDisSc = Trim(txtExtraDisSc.text)
RS!ExtraDisScDet = Trim(txtExtraDiscountDet.text)
RS!ExtraDisScNot = Trim(txtExtraSPNote.text)

RS!RetPolicy = Trim(txtRetPolicy.text)
RS!RetPolicyDet = Trim(txtRetPolicyDet.text)

RS!ExtraDisAlpha = Trim(txtExtraDisAlpha.text)
RS!RetPolicyAlpha = Trim(txtRetPolicyAlpha.text)
RS!GenTermINI = UCase(txtGeneral_Alpha.text)


RS.update


'cmdSave_2_Click
'Command1_Click
'Command3_Click

''MsgBox "Data Saved...", vbInformation

Exit Sub


save_:

MsgBox "" & err.DESCRIPTION

End Sub



'''Private Sub Command3_Click()
'''
''''On Error GoTo save_
'''
'''Set RS = New ADODB.Recordset
'''If RS.State = 1 Then RS.close
'''RS.Open "select * from AgreementMain where AgmNo=" & txtAgmNo & "", con, adOpenDynamic, adLockOptimistic
'''If RS.EOF = True Then
'''RS.AddNew
'''End If
'''
'''
'''
'''
'''RS!turnOverPart1 = vs_page2.TextMatrix(1, 0)
'''RS!turnOverDis1 = vs_page2.TextMatrix(1, 1)
'''RS!turnOverPart2 = vs_page2.TextMatrix(2, 0)
'''RS!turnOverDis2 = vs_page2.TextMatrix(2, 1)
'''RS!turnOverPart3 = vs_page2.TextMatrix(3, 0)
'''RS!turnOverDis3 = vs_page2.TextMatrix(3, 1)
'''RS!turnOverPart4 = vs_page2.TextMatrix(4, 0)
'''RS!turnOverDis4 = vs_page2.TextMatrix(4, 1)
'''RS!turnOverPart5 = vs_page2.TextMatrix(5, 0)
'''RS!turnOverDis5 = vs_page2.TextMatrix(5, 1)
'''RS!turnOverPart6 = vs_page2.TextMatrix(6, 0)
'''RS!turnOverDis6 = vs_page2.TextMatrix(6, 1)
'''RS!turnOver_SpNote = txtTurnOverSp.Text
'''RS!Transportation = txtTransportation
'''RS!TransportAlfa = txtTrans_Alfabat.Text
'''
''''---------------Cash Deposit==================================
'''
'''If Check1_cashdep.value = 1 Then
'''   RS!CashDepositYesNo = "n"
'''Else
'''   RS!CashDepositYesNo = "y"
'''End If
'''
'''RS!CashDepositSh = Trim(txtCashDepoSch.Text)
'''RS!CashDepositPeriod1 = vsCashDeposit.TextMatrix(1, 0)
'''RS!CashDepositExtraCr1 = vsCashDeposit.TextMatrix(1, 1)
'''RS!CashDepositPeriod2 = vsCashDeposit.TextMatrix(2, 0)
'''RS!CashDepositExtraCr2 = vsCashDeposit.TextMatrix(2, 1)
'''RS!CashDepositPeriod3 = vsCashDeposit.TextMatrix(3, 0)
'''RS!CashDepositExtraCr3 = vsCashDeposit.TextMatrix(3, 1)
'''RS!CashDepositPeriod4 = vsCashDeposit.TextMatrix(4, 0)
'''RS!CashDepositExtraCr4 = vsCashDeposit.TextMatrix(4, 1)
'''RS!CashDepositPeriod5 = vsCashDeposit.TextMatrix(5, 0)
'''RS!CashDepositExtraCr5 = vsCashDeposit.TextMatrix(5, 1)
'''
'''RS!CashDepositSpNot = Trim(txtCashDepSpNot.Text)
'''
'''
'''RS!CashDepositMinAmt1 = txtCashMinAmt1.Text
'''RS!CashDepositMinAmt2 = txtCashMinAmt2.Text
'''
'''RS!CashDepositMinExtra1 = txtCashMinExtra1.Text
'''RS!CashDepositMinExtra2 = txtCashMinExtra2.Text
'''
'''
''''---------------End Cash Deposit===============================
'''
'''RS!GenTerms_a = Trim(txtgenTerm_a.Text)
'''RS!GenTerms_b = Trim(txtgenTerm_b.Text)
'''RS!GenTerms_c = Trim(txtgenTerm_c.Text)
'''RS!GenTerms_d = Trim(txtgenTerm_d.Text)
'''RS!GenTerms_e = Trim(txtgenTerm_e.Text)
'''
''''---------------End Gen Terms==================================
'''
'''RS!turnOverINI = UCase(txtTurnDis_Alpha.Text)
'''RS!CashDepositINI = UCase(txtCashDis_Alpha.Text)
'''RS!GenTermINI = UCase(txtGeneral_Alpha.Text)
'''
'''RS!CashDepositSc = Trim(txtCashDepositSch.Text)
'''RS!CashDepositSch_Heading = Trim(txtCashDepositSch_Heading)
'''
'''
'''RS.update
'''
'''
'''MsgBox "Data Saled....", vbInformation
'''
'''Exit Sub
'''
'''save_:
'''
'''MsgBox "" & err.DESCRIPTION
'''
'''End Sub

Private Sub Command3_Click()

On Error GoTo save_

Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.close
RS.Open "select * from AgreementMain where AgmNo=" & txtAgmNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
RS.AddNew
End If

If Check1_turnOver.value = 0 Then
   RS!turnOver_yesno = "n"
Else
   RS!turnOver_yesno = "y"
End If

RS!turnOverDis = txtTurnOverDis.text


RS!turnOverPart1 = vs_page2.TextMatrix(1, 0)
RS!turnOverPart2 = vs_page2.TextMatrix(1, 1)
RS!turnOverPart3 = vs_page2.TextMatrix(1, 2)
RS!turnOverPart4 = vs_page2.TextMatrix(1, 3)
RS!turnOverPart5 = vs_page2.TextMatrix(1, 4)


'
'RS!turnOverPart1 = vs_page2.TextMatrix(1, 0)
'RS!turnOverDis1 = vs_page2.TextMatrix(1, 1)
'RS!turnOverPart2 = vs_page2.TextMatrix(2, 0)
'RS!turnOverDis2 = vs_page2.TextMatrix(2, 1)
'RS!turnOverPart3 = vs_page2.TextMatrix(3, 0)
'RS!turnOverDis3 = vs_page2.TextMatrix(3, 1)
'RS!turnOverPart4 = vs_page2.TextMatrix(4, 0)
'RS!turnOverDis4 = vs_page2.TextMatrix(4, 1)
'RS!turnOverPart5 = vs_page2.TextMatrix(5, 0)
'RS!turnOverDis5 = vs_page2.TextMatrix(5, 1)
'RS!turnOverPart6 = vs_page2.TextMatrix(6, 0)
'RS!turnOverDis6 = vs_page2.TextMatrix(6, 1)

RS!turnOver_SpNote = txtTurnOverSp.text
RS!turnOver_SpNote1 = txtTurnOverSp1.text
RS!turnOver_SpNote2 = txtTurnOverSp3.text

RS!Transportation = txtTransportation.text
RS!Transportation1 = txtTransportation1.text

RS!TransportAlfa = txtTrans_Alfabat.text

'---------------Cash Deposit==================================

If Check1_cashdep.value = 0 Then
   RS!CashDepositYesNo = "n"
Else
   RS!CashDepositYesNo = "y"
End If

RS!CashDepositSh = Trim(txtCashDepoSch.text)

RS!CashExamateSep1 = vsCashDeposit.TextMatrix(0, 1)
RS!CashExamateSep2 = vsCashDeposit.TextMatrix(0, 2)
RS!CashExamateSep3 = vsCashDeposit.TextMatrix(0, 3)
RS!CashExamateSep4 = vsCashDeposit.TextMatrix(0, 4)


RS!CashDepositPeriod1 = vsCashDeposit.TextMatrix(1, 0)
RS!CashDepositExtraCr1 = vsCashDeposit.TextMatrix(1, 1)

RS!CashDepositPeriod2 = vsCashDeposit.TextMatrix(1, 2)
RS!CashDepositExtraCr2 = vsCashDeposit.TextMatrix(1, 3)
RS!CashDepositPeriod3 = vsCashDeposit.TextMatrix(1, 4)

If Len(vsCashDeposit.TextMatrix(2, 4)) > 0 Then
   RS!CashDepositPeriod4 = vsCashDeposit.TextMatrix(2, 4)
Else
   RS!CashDepositPeriod4 = ""
End If


RS!CashDepositSpNot = Trim(txtCashDepSpNot.text)
RS!examdesc1 = txtCashDepRem1.text



RS!CashDepositMinAmt1 = txtCashMinAmt1.text
RS!CashDepositMinAmt2 = txtCashMinAmt2.text

RS!CashDepositMinExtra1 = txtCashMinExtra1.text
RS!CashDepositMinExtra2 = txtCashMinExtra2.text


'---------------End Cash Deposit===============================

RS!turnOverINI = UCase(txtTurnDis_Alpha.text)
RS!CashDepositINI = UCase(txtCashDis_Alpha.text)
RS!CashDepositSc = Trim(txtCashDepositSch.text)
RS!CashDepositSch_Heading = Trim(txtCashDepositSch_Heading)
RS!CashDepositAlpha = (txtCashDepositAlpha.text)

If Check1_TDExamte.value = 0 Then
   RS!TDExamate_y_n = "n"
Else
   RS!TDExamate_y_n = "1"
End If


RS!ExtraDisAlpha = Trim(txtExtraDisAlpha.text)
RS!RetPolicyAlpha = Trim(txtRetPolicyAlpha.text)
RS!GenTermINI = UCase(txtGeneral_Alpha.text)


RS!examdesc2 = Combo1_State.text

RS.update



'MsgBox "Data Saled....", vbInformation

Exit Sub

save_:

MsgBox "" & err.DESCRIPTION

End Sub

Private Sub Commanddelete_Click()

If txtAgmNo = "" Then
   MsgBox "Please Search Record...", vbInformation
   Exit Sub
End If


'======================================
 Dim rs1 As ADODB.Recordset
 Set rs1 = New ADODB.Recordset
 
 If rs1.State = 1 Then rs1.close
 rs1.Open "select top 1 * from AgreementMain where agmno=" & txtAgmNo.text & "", con, adOpenStatic, adLockReadOnly
 If rs1.EOF = False Then
     If rs1!bAuthorized = True Then
         MsgBox "You can'nt change, Already Locked !!", vbExclamation, "Alert"
         Exit Sub
     End If
    
 End If

'=======================================

If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "delete from AgreementMain where AgmNo=" & txtAgmNo & ""
   con.Execute "delete from AgreementDisStructure  where AgmNo=" & txtAgmNo & ""
   
   
End If

End Sub

Private Sub Commandedit_Click()

Edit = True
Commandedit.Enabled = False
cmdSave_2.Enabled = True
Commanddelete.Enabled = True
cmdSave_2.SetFocus


End Sub

Private Sub CommandPrint_Click()

DSNNew

CR.Reset
If Combo1_State.text = "ALL" Or Combo1_State.text = "" Then
   CR.ReportFileName = rptPath & "/agreement.rpt"
ElseIf Combo1_State.text = "TN" Then
   CR.ReportFileName = rptPath & "/agreementTN.rpt"
ElseIf Combo1_State.text = "AP" Then
   CR.ReportFileName = rptPath & "/agreementAP.rpt"
End If

If InStr(txtPName.text, ":") > 0 Then
   code_ = Mid(txtPName.text, 1, InStr(txtPName.text, ":") - 1)
End If



CR.ReplaceSelectionFormula "{AgreementMain.agmno}=" & txtAgmNo & ""
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
If rs1.State = 1 Then rs1.close
rs1.Open "select subledger from sledger where code='" & code_ & "'", con
If rs1.EOF = False Then
   CR.Formulas(5) = "add3='" & rs1(0) & "'"
Else
   vv1 = txtPName.text & "," & txtAddress2
   CR.Formulas(5) = "add3='" & vv1 & "'"
End If

CR.WindowShowPrintSetupBtn = True
CR.WindowShowPrintBtn = True
CR.WindowShowExportBtn = True
CR.WindowState = crptMaximized
CR.Action = 1

End Sub

Private Sub CommandReturn_Click()
Unload Me
End Sub
Sub drearText()

txtCashMinAmt1.text = "Rs. 10 Lacs"
txtCashMinAmt2.text = "Rs. 20 Lacs"
txtCashMinExtra1.text = "1 %"
txtCashMinExtra2.text = "2 %"

If Option1_new.value = True Then

txtwarm2.text = "We are pleased to have received your customer agreement form and appreciate the interest that you have shown in our books. We thank you for the courtesy and cooperation extended at your end. We strongly believe that you are one of the best entrepreneurs of your region and our business association will be beneficial for schools that are looking for quality books for their students."

If rs1.State = 1 Then rs1.close
rs1.Open "select TurnOverDiscount,turnOver_SpNote,turnOver_SpNote1,turnOver_SpNote2,TradeDisSp1_New,TradeDisSp2_New," & _
"TradeDisSp3_New,TransSp1,TransSp2,CashDepSp1,CashDepSp2, CashDepPeriod1," & _
"CashDepPeriod2,CashDepPeriod3,CashDepPeriod4,CashDepPeriod1v,CashDepPeriod2v,CashDepPeriod3v,CashDepPeriod4v,CashDepPeriod5v," & _
"CashDepMinAmt1,CashDepMinAmt2,CashDepMinPer1,CashDepMinPer2,ReturnPolicy,ReturnPolicy_a," & _
"GenTerms_a,GenTerms_b,GenTerms_c,GenTerms_d,GenTerms_e,GenTerms_f,GenTerms_g,GenTerms_h,GenTerms_i,GenTerms_j,GenTerms_k" & _
",AboveTerms from AgreementMainMaster", con
If rs1.EOF = False Then

    txtAboveBus.text = rs1!AboveTerms & ""
    txtgenTerm_a.text = rs1!GenTerms_a & ""
    txtgenTerm_b.text = rs1!GenTerms_b & ""
    txtgenTerm_c.text = rs1!GenTerms_c & ""
    txtgenTerm_d.text = rs1!GenTerms_d & ""
    txtgenTerm_e.text = rs1!GenTerms_e & ""
    txtgenTerm_f.text = rs1!GenTerms_f & ""
    txtgenTerm_g.text = rs1!GenTerms_g & ""
    txtgenTerm_h.text = rs1!GenTerms_h & ""
    txtgenTerm_i.text = rs1!GenTerms_i & ""
    txtgenTerm_j.text = rs1!GenTerms_j & ""
    txtgenTerm_k.text = rs1!GenTerms_k & ""
    
    
    
    
    
    txtRetPolicy.text = rs1!ReturnPolicy & ""
    
    txtRetPolicyDet.text = rs1!ReturnPolicy_a & ""

    'txtCashMinAmt1.text = rs1!CashDepMinAmt1 & ""
    'txtCashMinAmt2.text = rs1!CashDepMinAmt2 & ""
    
    'txtCashMinExtra1.text = rs1!CashDepMinPer1 & ""
    'txtCashMinExtra2.text = rs1!CashDepMinPer2 & ""
    
    
    
    txtTurnOverDis.text = rs1!TurnOverDiscount
    txtTurnOverSp.text = rs1!turnOver_SpNote
    txtTurnOverSp1.text = rs1!turnOver_SpNote1
    txtTurnOverSp3.text = rs1!turnOver_SpNote2

     txtSpNoteA.text = rs1!TradeDisSp1_New & ""
     'txtSpNoteB.text = rs1!TradeDisSp2_New & ""
     txtSpNoteC.text = rs1!TradeDisSp3_New & ""
     
     txtTransportation.text = rs1!TransSp1 & ""
     txtTransportation1 = rs1!TransSp2 & ""
     
     txtCashDepoSch.text = rs1!CashDepSp1
     txtCashDepSpNot.text = rs1!CashDepSp2
     
     
     
'    vsCashDeposit.TextMatrix(1, 0) = rs1!CashDepPeriod1
'    vsCashDeposit.TextMatrix(2, 0) = rs1!CashDepPeriod2
''    vsCashDeposit.TextMatrix(3, 0) = rs1!CashDepPeriod3
''    vsCashDeposit.TextMatrix(4, 0) = rs1!CashDepPeriod4
''    vsCashDeposit.TextMatrix(5, 0) = rs1!CashDepPeriod4
'
'    vsCashDeposit.TextMatrix(1, 1) = rs1!CashDepPeriod1v
'    vsCashDeposit.TextMatrix(2, 1) = rs1!CashDepPeriod2v
''    vsCashDeposit.TextMatrix(3, 1) = rs1!CashDepPeriod3v
'    vsCashDeposit.TextMatrix(4, 1) = rs1!CashDepPeriod4v
'    vsCashDeposit.TextMatrix(5, 1) = rs1!CashDepPeriod5v
    
    
    
     
     
End If

Else

txtwarm2.text = "Thank you for your support last year.  We value our partnership and believe your leadership will continue benefiting schools in need of quality books."

If rs1.State = 1 Then rs1.close
rs1.Open "select TurnOverDiscount,turnOver_SpNote,turnOver_SpNote1,turnOver_SpNote2,TradeDisSp1_Old,TradeDisSp2_Old," & _
"TradeDisSp3_Old,TransSp1,TransSp2,CashDepSp1,CashDepSp2,CashDepPeriod1," & _
"CashDepPeriod2,CashDepPeriod3,CashDepPeriod4,CashDepPeriod1v,CashDepPeriod2v,CashDepPeriod3v,CashDepPeriod4v,CashDepPeriod5v," & _
"CashDepMinAmt1,CashDepMinAmt2,CashDepMinPer1,CashDepMinPer2,ReturnPolicy,ReturnPolicy_a, " & _
"GenTerms_a,GenTerms_b,GenTerms_c,GenTerms_d,GenTerms_e,GenTerms_f,GenTerms_g,GenTerms_h,GenTerms_i,GenTerms_j,GenTerms_k" & _
",AboveTerms from AgreementMainMaster", con
If rs1.EOF = False Then


   txtAboveBus.text = rs1!AboveTerms & ""

   txtgenTerm_a.text = rs1!GenTerms_a & ""
    txtgenTerm_b.text = rs1!GenTerms_b & ""
    txtgenTerm_c.text = rs1!GenTerms_c & ""
    txtgenTerm_d.text = rs1!GenTerms_d & ""
    txtgenTerm_e.text = rs1!GenTerms_e & ""
    txtgenTerm_f.text = rs1!GenTerms_f & ""
    txtgenTerm_g.text = rs1!GenTerms_g & ""
    txtgenTerm_h.text = rs1!GenTerms_h & ""
    txtgenTerm_i.text = rs1!GenTerms_i & ""
    txtgenTerm_j.text = rs1!GenTerms_j & ""
    txtgenTerm_k.text = rs1!GenTerms_k & ""
    

    txtRetPolicy.text = rs1!ReturnPolicy & ""
    txtRetPolicyDet.text = rs1!ReturnPolicy_a & ""

     
    txtTurnOverDis.text = rs1!TurnOverDiscount
    txtTurnOverSp.text = rs1!turnOver_SpNote
    txtTurnOverSp1.text = rs1!turnOver_SpNote1
    txtTurnOverSp3.text = rs1!turnOver_SpNote2



     txtSpNoteA.text = rs1!TradeDisSp1_Old & ""
     'txtSpNoteB.text = rs1!TradeDisSp2_Old & ""
     txtSpNoteC.text = rs1!TradeDisSp3_Old & ""
     
     txtTransportation.text = rs1!TransSp1 & ""
     txtTransportation1 = rs1!TransSp2 & ""
     
     txtCashDepoSch.text = rs1!CashDepSp1
     txtCashDepSpNot.text = rs1!CashDepSp2

     txtCashDepRem1.text = rs1!CashDepPeriod3v
     
    vsCashDeposit.TextMatrix(1, 0) = rs1!CashDepPeriod1
    vsCashDeposit.TextMatrix(1, 1) = rs1!CashDepPeriod2
    vsCashDeposit.TextMatrix(1, 2) = rs1!CashDepPeriod3
    vsCashDeposit.TextMatrix(1, 3) = rs1!CashDepPeriod4
    vsCashDeposit.TextMatrix(1, 4) = rs1!CashDepPeriod1v
    vsCashDeposit.TextMatrix(2, 4) = rs1!CashDepPeriod2v
    
    
    
'    vsCashDeposit.TextMatrix(3, 0) = rs1!CashDepPeriod3
'    vsCashDeposit.TextMatrix(4, 0) = rs1!CashDepPeriod4
'    vsCashDeposit.TextMatrix(5, 0) = rs1!CashDepPeriod4
    
'    vsCashDeposit.TextMatrix(1, 1) = rs1!CashDepPeriod1v
'    vsCashDeposit.TextMatrix(2, 1) = rs1!CashDepPeriod2v
'    vsCashDeposit.TextMatrix(3, 1) = rs1!CashDepPeriod3v
'    vsCashDeposit.TextMatrix(4, 1) = rs1!CashDepPeriod4v
'    vsCashDeposit.TextMatrix(5, 1) = rs1!CashDepPeriod5v
    
    
 
     
End If


End If

End Sub
Sub fetchStateWiseCD()

Dim rss4 As New ADODB.Recordset

Set rss4 = New ADODB.Recordset

vsCashDeposit.Clear

If rss4.State = 1 Then rss4.close
rss4.Open "select CashDepPeriod1,CashDepPeriod2,CashDepPeriod3,CashDepPeriod4,CashDepPeriod1v,CashDepPeriod2v" & _
",CashExamateAmt1,CashExamateAmt2,CashExamateAmt3,CashExamateAmt4,CashExamateAmt5,CashExamateMay1,CashExamateMay2,CashExamateMay3,CashExamateMay4,CashDepSp2 from AgreementMainMaster ", con
If rss4.EOF = False Then

    If (Combo1_State.text = "ALL") Then
    
        vsCashDeposit.TextMatrix(1, 0) = rss4!CashDepPeriod1
        vsCashDeposit.TextMatrix(1, 1) = rss4!CashDepPeriod2
        vsCashDeposit.TextMatrix(1, 2) = rss4!CashDepPeriod3
        vsCashDeposit.TextMatrix(1, 3) = rss4!CashDepPeriod4
        vsCashDeposit.TextMatrix(1, 4) = rss4!CashDepPeriod1v
        vsCashDeposit.TextMatrix(2, 4) = rss4!CashDepPeriod2v
        
        vsCashDeposit.TextMatrix(0, 0) = "Period"
        vsCashDeposit.TextMatrix(0, 1) = "January"
        vsCashDeposit.TextMatrix(0, 2) = "February"
        vsCashDeposit.TextMatrix(0, 3) = "March"
        vsCashDeposit.TextMatrix(0, 4) = "April"
        
        txtCashDepSpNot.text = rss4!CashDepSp2 & ""
    
    ElseIf (Combo1_State.text = "TN") Then
    
        vsCashDeposit.TextMatrix(1, 0) = rss4!CashExamateAmt1
        vsCashDeposit.TextMatrix(1, 1) = rss4!CashExamateAmt2
        vsCashDeposit.TextMatrix(1, 2) = rss4!CashExamateAmt3
        vsCashDeposit.TextMatrix(1, 3) = rss4!CashExamateAmt4
        vsCashDeposit.TextMatrix(1, 4) = rss4!CashExamateAmt5
        'vsCashDeposit.TextMatrix(2, 4) = rss4!CashExamateAmt6
        
        vsCashDeposit.TextMatrix(0, 0) = "Period"
        vsCashDeposit.TextMatrix(0, 1) = rss4!CashExamateMay1
        vsCashDeposit.TextMatrix(0, 2) = rss4!CashExamateMay2
        vsCashDeposit.TextMatrix(0, 3) = rss4!CashExamateMay3
        vsCashDeposit.TextMatrix(0, 4) = rss4!CashExamateMay4

        txtCashDepSpNot.text = "Customers depositing the following amounts in a single payment between 01-04-2025 and 31-07-2025 will receive an additional percentage, as per the table below:"
    
       ElseIf (Combo1_State.text = "AP") Then
    
        vsCashDeposit.TextMatrix(1, 0) = "CD %"
        vsCashDeposit.TextMatrix(1, 1) = "5%  "
        vsCashDeposit.TextMatrix(1, 2) = "3%  "
        
        vsCashDeposit.TextMatrix(0, 0) = "Period"
        vsCashDeposit.TextMatrix(0, 1) = "JUNE"
        vsCashDeposit.TextMatrix(0, 2) = "JULY"
     
        txtCashDepSpNot.text = "Customers depositing the following amounts in a single payment between 01-06-2025 and 31-07-2025 will receive an additional percentage, as per the table below:"
    
    
    End If

End If



vsCashDeposit.ColWidth(0) = 2200
vsCashDeposit.ColWidth(1) = 2200
vsCashDeposit.ColWidth(2) = 2200
vsCashDeposit.ColWidth(3) = 2200
vsCashDeposit.ColWidth(4) = 2200
End Sub
Sub iniForm()
txtDate.value = Format(Date, "dd/MM/yyyy")

Edit = False



drearText


fillGrid

Dim warm_, subject_ As String

warm_ = "As per our mutual discussion, we are pleased to share the following business terms for the session"

subject_ = "MOU for the year"

txtSub_yrs.text = session
txtSession1.text = session

txtSubject.text = subject_ & " " & session
txtwarm3.text = warm_ & "  " & session

vs.TextMatrix(1, 0) = "1"
vs.TextMatrix(2, 0) = "2"
vs.TextMatrix(3, 0) = "3"
vs.TextMatrix(4, 0) = "4"
vs.TextMatrix(5, 0) = "5"
vs.TextMatrix(6, 0) = "6"
vs.TextMatrix(7, 0) = "7"
vs.TextMatrix(8, 0) = "8"


kk1 = 1
vs.rows = 1
If rs1.State = 1 Then rs1.close
rs1.Open "select [SN],description,[Class],[Percent],[DiscountOffer] from AgmDisStructure_Master order by sn"
While rs1.EOF = False
vs.rows = vs.rows + 1
vs.TextMatrix(kk1, 0) = rs1.Fields("SN").value
vs.TextMatrix(kk1, 1) = rs1.Fields("description").value
vs.TextMatrix(kk1, 2) = rs1.Fields("Class").value
vs.TextMatrix(kk1, 3) = rs1.Fields("Percent").value
vs.TextMatrix(kk1, 4) = rs1.Fields("DiscountOffer").value & ""
kk1 = kk1 + 1

rs1.MoveNext
Wend



vs_page2.TextMatrix(0, 0) = "5 Lac to <15 Lac"
vs_page2.TextMatrix(0, 1) = "15 Lac to< 25 Lac"
vs_page2.TextMatrix(0, 2) = "25 Lac to< 50 Lac"
vs_page2.TextMatrix(0, 3) = "50 Lac to< 1 Crore"
vs_page2.TextMatrix(0, 4) = "1 Crore and Above"


vs_page2.TextMatrix(1, 0) = "4 %"
vs_page2.TextMatrix(1, 1) = "5 %"
vs_page2.TextMatrix(1, 2) = "6 %"
vs_page2.TextMatrix(1, 3) = "7 %"
vs_page2.TextMatrix(1, 4) = "8 %"


vs_page2.ColWidth(0) = 2200
vs_page2.ColWidth(1) = 2200
vs_page2.ColWidth(2) = 2200
vs_page2.ColWidth(3) = 2200
vs_page2.ColWidth(4) = 2200



vsCashDeposit.TextMatrix(0, 0) = "Period"
vsCashDeposit.TextMatrix(0, 1) = "January"
vsCashDeposit.TextMatrix(0, 2) = "February"
vsCashDeposit.TextMatrix(0, 3) = "March"
vsCashDeposit.TextMatrix(0, 4) = "April"

vsCashDeposit.ColWidth(0) = 2200
vsCashDeposit.ColWidth(1) = 2200
vsCashDeposit.ColWidth(2) = 2200
vsCashDeposit.ColWidth(3) = 2200
vsCashDeposit.ColWidth(4) = 2200


VSExamtion.TextMatrix(0, 0) = "Description"
VSExamtion.TextMatrix(0, 1) = "Class"
VSExamtion.TextMatrix(0, 2) = "Percent"




VSExamtion.ColWidth(0) = 6500
VSExamtion.ColWidth(1) = 1700
VSExamtion.ColWidth(2) = 1700




txtgenTerm_l.text = "Bank Name: ICICI BANK LTD"
txtgenTerm_m.text = "Account No: 628505018973,  IFS Code: ICICI0006285"



txtbName.text = "Account Name: BLUEPRINT EDUCATION (A DIVISION OF CHITRA PRAKASHAN INDIA PVT LTD)"


txtAgmNo = MaxSNo("AgreementMain", "agmno")

Combo1_State.ListIndex = 0


End Sub
Private Sub Form_Load()

Me.Height = 11000
Me.Width = 16980
Me.top = 1
Me.Left = 50


iniForm

End Sub
Sub fillGrid()

vs.FormatString = "SN.|Description|Class|Trade Discount|Discount Offered"
vs.ColWidth(0) = 600
vs.ColWidth(1) = 4500
vs.ColWidth(2) = 2000
vs.ColWidth(3) = 2000
vs.ColWidth(4) = 1500

End Sub

Private Sub Option1_new_Click()
drearText
End Sub

Private Sub Option2_Existing_Click()
drearText
End Sub

Private Sub txtAgmNo_GotFocus()
      
      If PopUpValue1 <> "" Then
         txtAgmNo = PopUpValue1
         searchData
         
         PopUpValue1 = ""
         PopUpValue2 = ""
      End If
      
End Sub

Private Sub txtAgmNo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then
      popuplist1 "Select distinct AgmNo,PName from AgreementMain order by AgmNo", con
   End If
   
   If KeyCode = 13 Then
       searchData
   End If
   
End Sub
Private Sub txtName_GotFocus()

If PopUpValue1 <> "" Then
   txtPName.text = PopUpValue3 & ":" & PopUpValue1
   'txtaddress1.text = PopUpValue2
   
   If RS.State = 1 Then RS.close
   RS.Open "select a.DISTCODE,a.Pin,a.phone,a.mobile,a.email,a.contactp,a.ADDRESS1,a.ADDRESS2,a.states,b.city,b.District from sledger as a inner join PartyDetailQry as b on (a.code=b.code) where a.code='" & PopUpValue3 & "'", con
   If RS.EOF = False Then
      
      If RS!address1 <> "" Then
         txtAddress1.text = RS!address1 & IIf(RS!address2 = "", "", "," & RS!address2)
      End If
      
      If (RS!distcode = RS!city) Then
        txtAddress2.text = RS!distcode
      Else
        txtAddress2.text = RS!city & "(" & RS!distcode & ")"
      End If
      
      If RS!pin <> "" Then
         If (RS!distcode = RS!city) Then
            txtAddress2.text = RS!distcode & " - " & RS!pin
         Else
            txtAddress2.text = RS!city & "(" & RS!distcode & ")" & " - " & RS!pin
         End If
      End If
      
      txtAddress3.text = RS!states & ""
      
      txtAddress4.text = "PH.-" & RS!phone & ""
      
      txtMobile.text = RS!mobile & ""
      txtEmail.text = RS!email & ""
      txtName.text = RS!contactp & ""
      
      'If RS!address1 <> "" Then
      '   txtaddress1.text = RS!address1
      'End If
      
      
      
      
      
   End If
   'txtAddress3.text = PopUpValue3
   
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
End If

End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    searchType = "party"
    popuplist_client "Select DESCFORINVOICE as Party,City,Code from PartyDetailQry where gledger='SUNDRY DEBTORS' and " & stringyear & " order by DESCFORINVOICE", con

End If
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
        
        If KeyCode = 115 Then
           vs.RemoveItem (vs.RowSel)
           
           For k1 = 1 To vs.rows - 1
               vs.TextMatrix(k1, 0) = k1
           Next
        End If
        
        
        
        If KeyCode = 13 Then
           vs.TextMatrix(vs.RowSel, 0) = vs.Row
           
           If (vs.Col = 3) Then
              sendkeys "{down}"
           ElseIf (vs.Col = 4) Then
              sendkeys "{down}"
           End If
           
           If (vs.Col >= 3 And vs.Col <= 4) Then
            kk2 = vs.TextMatrix(vs.RowSel, vs.Col)
            
            If (Len(kk2) > 0 And Len(kk2) <= 2) Then
                hh2 = InStr(kk2, "%")
                If hh2 = 0 Then
                    vs.TextMatrix(vs.RowSel, vs.Col) = vs.TextMatrix(vs.RowSel, vs.Col) & " %"
                End If
            End If
            
          End If
           
           
        End If
        
End Sub
