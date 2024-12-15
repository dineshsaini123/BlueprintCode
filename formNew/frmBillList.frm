VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillList 
   Caption         =   "Chitra"
   ClientHeight    =   10212
   ClientLeft      =   492
   ClientTop       =   780
   ClientWidth     =   17412
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10212
   ScaleWidth      =   17412
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10368
      Top             =   9936
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame panel 
      Height          =   9900
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   17376
      Begin VB.Frame f1 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   180
         TabIndex        =   77
         Top             =   330
         Visible         =   0   'False
         Width           =   90
         Begin VB.OptionButton party 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Party Wise Dr/Cr Entry"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3450
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   45
            Width           =   3360
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Bill Aouthorized Option"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   -120
            Width           =   3375
         End
         Begin VB.Image Image1 
            Height          =   675
            Left            =   15
            Stretch         =   -1  'True
            Top             =   60
            Width           =   10155
         End
      End
      Begin TabDlg.SSTab Opening 
         Height          =   9552
         Left            =   72
         TabIndex        =   3
         Top             =   180
         Width           =   17280
         _ExtentX        =   30480
         _ExtentY        =   16849
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         BackColor       =   12058623
         TabCaption(0)   =   "Bill Authorized Option"
         TabPicture(0)   =   "frmBillList.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label2"
         Tab(0).Control(1)=   "Label1(0)"
         Tab(0).Control(2)=   "vs"
         Tab(0).Control(3)=   "FromDate"
         Tab(0).Control(4)=   "toDate"
         Tab(0).Control(5)=   "bill"
         Tab(0).Control(6)=   "cmdset"
         Tab(0).Control(7)=   "Frame1"
         Tab(0).Control(8)=   "Frame2"
         Tab(0).Control(9)=   "pass"
         Tab(0).Control(10)=   "cmdPrint_Pro"
         Tab(0).Control(11)=   "vs_promotion"
         Tab(0).Control(12)=   "Command5_print"
         Tab(0).Control(13)=   "cmdUpDatePromotion"
         Tab(0).Control(14)=   "cmdRefProm"
         Tab(0).Control(15)=   "frmPassword"
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "Dr/Cr Entry"
         TabPicture(1)   =   "frmBillList.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblCR"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "drLebel"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "CrLebel"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label20"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label19"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "phone"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label18"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label15"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label13"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label12"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label10"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label8"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label6"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Label5"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Label9"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "Label11"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Label4"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "lblTOD"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "Label16"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "Label1(21)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "Label1(20)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "lblTotalRecord"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "Label1(2)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "Label1(1)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "lblfrt"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "Label23"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "lblCAF(1)"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "VS_sale"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "vs_document"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "vs1"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "RecDates"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "crpt"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).Control(32)=   "selectAll"
         Tab(1).Control(32).Enabled=   0   'False
         Tab(1).Control(33)=   "cmdprintalf"
         Tab(1).Control(33).Enabled=   0   'False
         Tab(1).Control(34)=   "cmdModify"
         Tab(1).Control(34).Enabled=   0   'False
         Tab(1).Control(35)=   "cmdDel"
         Tab(1).Control(35).Enabled=   0   'False
         Tab(1).Control(36)=   "cmdRefresh"
         Tab(1).Control(36).Enabled=   0   'False
         Tab(1).Control(37)=   "cmdSave"
         Tab(1).Control(37).Enabled=   0   'False
         Tab(1).Control(38)=   "cmdMain"
         Tab(1).Control(38).Enabled=   0   'False
         Tab(1).Control(39)=   "txtalfa"
         Tab(1).Control(39).Enabled=   0   'False
         Tab(1).Control(40)=   "cboStation"
         Tab(1).Control(40).Enabled=   0   'False
         Tab(1).Control(41)=   "closingcr"
         Tab(1).Control(41).Enabled=   0   'False
         Tab(1).Control(42)=   "txtRem"
         Tab(1).Control(42).Enabled=   0   'False
         Tab(1).Control(43)=   "txtBalance"
         Tab(1).Control(43).Enabled=   0   'False
         Tab(1).Control(44)=   "cmdShow1"
         Tab(1).Control(44).Enabled=   0   'False
         Tab(1).Control(45)=   "Check1"
         Tab(1).Control(45).Enabled=   0   'False
         Tab(1).Control(46)=   "cboop"
         Tab(1).Control(46).Enabled=   0   'False
         Tab(1).Control(47)=   "txtOp"
         Tab(1).Control(47).Enabled=   0   'False
         Tab(1).Control(48)=   "cboParty"
         Tab(1).Control(48).Enabled=   0   'False
         Tab(1).Control(49)=   "txtRecno"
         Tab(1).Control(49).Enabled=   0   'False
         Tab(1).Control(50)=   "txtQty"
         Tab(1).Control(50).Enabled=   0   'False
         Tab(1).Control(51)=   "Receive"
         Tab(1).Control(51).Enabled=   0   'False
         Tab(1).Control(52)=   "Issue"
         Tab(1).Control(52).Enabled=   0   'False
         Tab(1).Control(53)=   "txtdes"
         Tab(1).Control(53).Enabled=   0   'False
         Tab(1).Control(54)=   "cmdOrderList"
         Tab(1).Control(54).Enabled=   0   'False
         Tab(1).Control(55)=   "frm"
         Tab(1).Control(55).Enabled=   0   'False
         Tab(1).Control(56)=   "Command3"
         Tab(1).Control(56).Enabled=   0   'False
         Tab(1).Control(57)=   "cmdPrint"
         Tab(1).Control(57).Enabled=   0   'False
         Tab(1).Control(58)=   "Check_rep_billwise"
         Tab(1).Control(58).Enabled=   0   'False
         Tab(1).Control(59)=   "txt_ason"
         Tab(1).Control(59).Enabled=   0   'False
         Tab(1).Control(60)=   "Timer1"
         Tab(1).Control(60).Enabled=   0   'False
         Tab(1).Control(61)=   "cmdNewPrint"
         Tab(1).Control(61).Enabled=   0   'False
         Tab(1).Control(62)=   "Command5"
         Tab(1).Control(62).Enabled=   0   'False
         Tab(1).Control(63)=   "Check3_newledger"
         Tab(1).Control(63).Enabled=   0   'False
         Tab(1).Control(64)=   "cmdTitleLedger"
         Tab(1).Control(64).Enabled=   0   'False
         Tab(1).Control(65)=   "txtClosing"
         Tab(1).Control(65).Enabled=   0   'False
         Tab(1).Control(66)=   "txtcr"
         Tab(1).Control(66).Enabled=   0   'False
         Tab(1).Control(67)=   "cmdBilty"
         Tab(1).Control(67).Enabled=   0   'False
         Tab(1).Control(68)=   "cboPartyList"
         Tab(1).Control(68).Enabled=   0   'False
         Tab(1).Control(69)=   "Timer2"
         Tab(1).Control(69).Enabled=   0   'False
         Tab(1).Control(70)=   "frmOrderList"
         Tab(1).Control(70).Enabled=   0   'False
         Tab(1).Control(71)=   "cmdDocument"
         Tab(1).Control(71).Enabled=   0   'False
         Tab(1).Control(72)=   "cmdMail_1"
         Tab(1).Control(72).Enabled=   0   'False
         Tab(1).Control(73)=   "Command7"
         Tab(1).Control(73).Enabled=   0   'False
         Tab(1).Control(74)=   "Command6"
         Tab(1).Control(74).Enabled=   0   'False
         Tab(1).Control(75)=   "Command8print"
         Tab(1).Control(75).Enabled=   0   'False
         Tab(1).Control(76)=   "txtchequeeNo"
         Tab(1).Control(76).Enabled=   0   'False
         Tab(1).Control(77)=   "List1_ch"
         Tab(1).Control(77).Enabled=   0   'False
         Tab(1).ControlCount=   78
         TabCaption(2)   =   "Opening"
         TabPicture(2)   =   "frmBillList.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblStation"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label21"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label17"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "abc"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Label22"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Shape1"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Shape2"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "Shape3"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Shape4"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "date2"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "date1"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "vsop"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "comdio"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "dateason"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "Check_ClosingDesc"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "Command4"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).Control(16)=   "Frame3"
         Tab(2).Control(16).Enabled=   0   'False
         Tab(2).Control(17)=   "Check2"
         Tab(2).Control(17).Enabled=   0   'False
         Tab(2).Control(18)=   "cmdAson"
         Tab(2).Control(18).Enabled=   0   'False
         Tab(2).Control(19)=   "dataTrans"
         Tab(2).Control(19).Enabled=   0   'False
         Tab(2).Control(20)=   "Command2"
         Tab(2).Control(20).Enabled=   0   'False
         Tab(2).Control(21)=   "cboStation1"
         Tab(2).Control(21).Enabled=   0   'False
         Tab(2).Control(22)=   "txtamount"
         Tab(2).Control(22).Enabled=   0   'False
         Tab(2).Control(23)=   "Command1"
         Tab(2).Control(23).Enabled=   0   'False
         Tab(2).Control(24)=   "cmdShowClosing"
         Tab(2).Control(24).Enabled=   0   'False
         Tab(2).Control(25)=   "COMBOGENLEDGER"
         Tab(2).Control(25).Enabled=   0   'False
         Tab(2).Control(26)=   "cmdupdatep"
         Tab(2).Control(26).Enabled=   0   'False
         Tab(2).Control(27)=   "Check3_ledgerClosingTrans"
         Tab(2).Control(27).Enabled=   0   'False
         Tab(2).Control(28)=   "cmdPrint1"
         Tab(2).Control(28).Enabled=   0   'False
         Tab(2).Control(29)=   "Check3_filter"
         Tab(2).Control(29).Enabled=   0   'False
         Tab(2).Control(30)=   "cmdRepQty"
         Tab(2).Control(30).Enabled=   0   'False
         Tab(2).ControlCount=   31
         Begin VB.ListBox List1_ch 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   984
            Left            =   11412
            TabIndex        =   152
            Top             =   2304
            Visible         =   0   'False
            Width           =   5808
         End
         Begin VB.CommandButton cmdRepQty 
            BackColor       =   &H00FAEFC9&
            Caption         =   "Print Excel"
            Height          =   408
            Left            =   -63372
            Style           =   1  'Graphical
            TabIndex        =   149
            Top             =   4716
            Width           =   2004
         End
         Begin VB.CheckBox Check3_filter 
            Caption         =   "A/c Not Match"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   -63300
            TabIndex        =   148
            Top             =   2052
            Width           =   2184
         End
         Begin VB.CommandButton cmdPrint1 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Print"
            Height          =   396
            Left            =   -63372
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   4284
            Width           =   2010
         End
         Begin VB.TextBox txtchequeeNo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   13284
            MaxLength       =   50
            TabIndex        =   146
            Top             =   1944
            Width           =   1488
         End
         Begin VB.CommandButton Command8print 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1764
            Style           =   1  'Graphical
            TabIndex        =   144
            Top             =   2880
            Width           =   888
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Create Pdf"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   456
            Left            =   13284
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   816
            Width           =   1488
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&View Document"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   168
            Left            =   13068
            Style           =   1  'Graphical
            TabIndex        =   143
            Top             =   1116
            Visible         =   0   'False
            Width           =   1488
         End
         Begin VB.CommandButton cmdMail_1 
            BackColor       =   &H00FAEFC9&
            Caption         =   "Send Mail"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3528
            Style           =   1  'Graphical
            TabIndex        =   142
            Top             =   2865
            Width           =   852
         End
         Begin VB.CommandButton cmdDocument 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Document Link"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   456
            Left            =   13284
            Style           =   1  'Graphical
            TabIndex        =   141
            Top             =   324
            Width           =   1488
         End
         Begin VB.Frame frmOrderList 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Order List"
            Height          =   5355
            Left            =   48
            TabIndex        =   102
            Top             =   3480
            Visible         =   0   'False
            Width           =   12144
            Begin VB.CommandButton cmdexit_1 
               Caption         =   "E&xit"
               Height          =   384
               Left            =   10872
               TabIndex        =   104
               Top             =   144
               Width           =   1200
            End
            Begin VSFlex7Ctl.VSFlexGrid VsOrderList 
               Height          =   4704
               Left            =   60
               TabIndex        =   103
               Top             =   540
               Width           =   12012
               _cx             =   21188
               _cy             =   8297
               _ConvInfo       =   1
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   16761024
               ForeColorFixed  =   -2147483630
               BackColorSel    =   16777215
               ForeColorSel    =   0
               BackColorBkg    =   16777215
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   2
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   200
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
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
               Editable        =   0
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
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   12060
            Top             =   324
         End
         Begin VB.ListBox cboPartyList 
            Appearance      =   0  'Flat
            Height          =   2616
            Left            =   11508
            Style           =   1  'Checkbox
            TabIndex        =   18
            Top             =   2784
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.CommandButton cmdBilty 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Bilty Details"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   9648
            Style           =   1  'Graphical
            TabIndex        =   138
            Top             =   2865
            Width           =   924
         End
         Begin VB.TextBox txtcr 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   7308
            Locked          =   -1  'True
            TabIndex        =   137
            Top             =   8895
            Width           =   1320
         End
         Begin VB.TextBox txtClosing 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   8895
            Width           =   1284
         End
         Begin VB.CommandButton cmdTitleLedger 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Title Ledger"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8784
            Style           =   1  'Graphical
            TabIndex        =   133
            Top             =   2865
            Width           =   855
         End
         Begin VB.Frame frmPassword 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enter Password and Press Enter"
            Height          =   870
            Left            =   -74910
            TabIndex        =   131
            Top             =   1890
            Visible         =   0   'False
            Width           =   2535
            Begin VB.TextBox txtEnterPass 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   495
               PasswordChar    =   "#"
               TabIndex        =   132
               Top             =   360
               Width           =   1725
            End
         End
         Begin VB.CommandButton cmdRefProm 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Refresh Sponsorship"
            Height          =   480
            Left            =   -64515
            Style           =   1  'Graphical
            TabIndex        =   130
            Top             =   1440
            Width           =   1470
         End
         Begin VB.CommandButton cmdUpDatePromotion 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Update Sponsorship"
            Height          =   480
            Left            =   -63030
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   1440
            Width           =   1470
         End
         Begin VB.CheckBox Check3_newledger 
            Caption         =   "For Ledger View(New && fast)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6210
            TabIndex        =   127
            Top             =   312
            Value           =   1  'Checked
            Width           =   2130
         End
         Begin VB.CommandButton Command5 
            Caption         =   "New Fatch Ledger"
            Height          =   420
            Left            =   4608
            TabIndex        =   126
            Top             =   1764
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.CommandButton cmdNewPrint 
            BackColor       =   &H00FAEFC9&
            Caption         =   "Print (&Outstanding)"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   564
            Left            =   13284
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   1284
            Width           =   1488
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   12420
            Top             =   360
         End
         Begin VB.CommandButton Command5_print 
            BackColor       =   &H00FFFFC0&
            Caption         =   "&Print"
            Height          =   435
            Left            =   -71460
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   1440
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VSFlex7Ctl.VSFlexGrid vs_promotion 
            Height          =   7320
            Left            =   -74880
            TabIndex        =   119
            Top             =   2595
            Visible         =   0   'False
            Width           =   14355
            _cx             =   25321
            _cy             =   12912
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16777209
            ForeColor       =   16711680
            BackColorFixed  =   16777173
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777166
            BackColorAlternate=   16777209
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmBillList.frx":0054
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
            ExplorerBar     =   7
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
         Begin VB.CommandButton cmdPrint_Pro 
            BackColor       =   &H00FFFFC0&
            Caption         =   "&View"
            Height          =   435
            Left            =   -72540
            Style           =   1  'Graphical
            TabIndex        =   118
            Top             =   1440
            Visible         =   0   'False
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker txt_ason 
            Height          =   315
            Left            =   13125
            TabIndex        =   116
            Top             =   3075
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   572
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   127991811
            CurrentDate     =   39979
         End
         Begin VB.CheckBox Check_rep_billwise 
            Caption         =   "Bill Wise"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   12180
            TabIndex        =   113
            Top             =   2760
            Width           =   1395
         End
         Begin VB.CheckBox Check3_ledgerClosingTrans 
            Caption         =   "Ledger Closing Transfar"
            Height          =   495
            Left            =   -63624
            TabIndex        =   112
            Top             =   8172
            Width           =   1575
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00FAEFC9&
            Caption         =   "Print PD&F"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2652
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   2865
            Width           =   888
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Ledger Dist. Wise"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4404
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   2865
            Width           =   900
         End
         Begin VB.Frame frm 
            Caption         =   "State/Rep/District Wise"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   756
            Left            =   8328
            TabIndex        =   105
            Top             =   1692
            Width           =   4920
            Begin VB.OptionButton Option2_mzn 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Manager Wise"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   384
               Left            =   2520
               TabIndex        =   134
               Top             =   276
               Width           =   1110
            End
            Begin VB.CheckBox Check3_AmountTobecollect 
               Caption         =   "Balance Amt. Only"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   516
               Left            =   3732
               TabIndex        =   111
               Top             =   144
               Value           =   1  'Checked
               Width           =   1164
            End
            Begin VB.OptionButton Option4_rep 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Rep. Wise"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   384
               Left            =   1545
               TabIndex        =   108
               Top             =   276
               Width           =   975
            End
            Begin VB.OptionButton Option3_dist 
               BackColor       =   &H00FFFFFF&
               Caption         =   "District"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   384
               Left            =   75
               TabIndex        =   107
               Top             =   276
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton Option2_state 
               BackColor       =   &H00FFFFFF&
               Caption         =   "State"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   384
               Left            =   930
               TabIndex        =   106
               Top             =   276
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdOrderList 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Order List"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8004
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2865
            Width           =   792
         End
         Begin VB.Frame pass 
            Height          =   465
            Left            =   -67785
            TabIndex        =   91
            Top             =   1455
            Visible         =   0   'False
            Width           =   2730
            Begin VB.TextBox txtadmin 
               Appearance      =   0  'Flat
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   90
               PasswordChar    =   "*"
               TabIndex        =   92
               Top             =   135
               Width           =   1590
            End
            Begin VB.Label Label14 
               Caption         =   "Press Enter"
               Height          =   165
               Left            =   1770
               TabIndex        =   93
               Top             =   180
               Width           =   900
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1050
            Left            =   -63072
            TabIndex        =   87
            Top             =   345
            Width           =   1515
            Begin VB.OptionButton All 
               Caption         =   "All"
               Height          =   255
               Left            =   45
               TabIndex        =   90
               Top             =   720
               Width           =   1290
            End
            Begin VB.OptionButton Unautho 
               Caption         =   "Un Authorized"
               Height          =   270
               Left            =   45
               TabIndex        =   89
               Top             =   435
               Value           =   -1  'True
               Width           =   1320
            End
            Begin VB.OptionButton autho 
               Caption         =   "Authorized"
               Height          =   180
               Left            =   45
               TabIndex        =   88
               Top             =   195
               Width           =   1095
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1110
            Left            =   -72555
            TabIndex        =   81
            Top             =   315
            Width           =   9144
            Begin VB.OptionButton Option2_agm 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Cust. Agreement"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   7236
               TabIndex        =   153
               Top             =   252
               Width           =   1632
            End
            Begin VB.OptionButton Option2_app 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Approval"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1305
               TabIndex        =   124
               Top             =   675
               Width           =   1200
            End
            Begin VB.OptionButton Option2_donation 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Promotion"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   114
               Top             =   675
               Width           =   1200
            End
            Begin VB.OptionButton Option_bookRetSp 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Book Return  (Specimen)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4815
               TabIndex        =   101
               Top             =   690
               Width           =   2295
            End
            Begin VB.OptionButton Option_bookIssueSp 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Book Issue  (Specimen)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2595
               TabIndex        =   100
               Top             =   690
               Width           =   2205
            End
            Begin VB.OptionButton dbit 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Debit Note"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   5925
               TabIndex        =   86
               Top             =   270
               Width           =   1200
            End
            Begin VB.OptionButton crdit 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Credit Note"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4485
               TabIndex        =   85
               Top             =   270
               Width           =   1200
            End
            Begin VB.OptionButton sales 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Sales Bill "
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Value           =   -1  'True
               Width           =   1380
            End
            Begin VB.OptionButton cash 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Cash Bill"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1530
               TabIndex        =   83
               Top             =   270
               Width           =   1200
            End
            Begin VB.OptionButton credit 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Credit Note Item"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2790
               TabIndex        =   82
               Top             =   270
               Width           =   1620
            End
         End
         Begin VB.CommandButton cmdset 
            BackColor       =   &H00FFFFC0&
            Caption         =   "S&ave"
            Height          =   435
            Left            =   -73920
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   1320
            Width           =   1260
         End
         Begin VB.Frame bill 
            Height          =   300
            Left            =   -74880
            TabIndex        =   47
            Top             =   9240
            Visible         =   0   'False
            Width           =   10335
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   720
               TabIndex        =   99
               Text            =   "Text1"
               Top             =   240
               Width           =   195
            End
            Begin VB.CommandButton cmdshow 
               BackColor       =   &H00FFFFC0&
               Caption         =   "&Show"
               Height          =   300
               Left            =   330
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   990
               Visible         =   0   'False
               Width           =   75
            End
            Begin VB.TextBox txtParty 
               Height          =   300
               Left            =   180
               TabIndex        =   48
               Top             =   990
               Visible         =   0   'False
               Width           =   135
            End
            Begin VB.Label Label7 
               Caption         =   "F2 For Search Party"
               Height          =   270
               Left            =   195
               TabIndex        =   51
               Top             =   1020
               Visible         =   0   'False
               Width           =   120
            End
            Begin VB.Label Label3 
               Caption         =   "Party"
               Height          =   225
               Left            =   195
               TabIndex        =   50
               Top             =   1035
               Visible         =   0   'False
               Width           =   90
            End
         End
         Begin VB.TextBox txtdes 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   46
            Top             =   1440
            Width           =   4764
         End
         Begin VB.OptionButton Issue 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Cr"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2532
            TabIndex        =   45
            Top             =   2388
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton Receive 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Dr"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3204
            TabIndex        =   44
            Top             =   2388
            Width           =   630
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   1080
            MaxLength       =   12
            TabIndex        =   43
            Top             =   2340
            Width           =   1380
         End
         Begin VB.TextBox txtRecno 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   1080
            TabIndex        =   0
            Top             =   660
            Width           =   1335
         End
         Begin VB.ComboBox cboParty 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   288
            Left            =   1080
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   42
            Top             =   1020
            Width           =   4764
         End
         Begin VB.TextBox txtOp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   11388
            MaxLength       =   50
            TabIndex        =   41
            Top             =   780
            Width           =   1365
         End
         Begin VB.ComboBox cboop 
            Height          =   288
            Left            =   12780
            Style           =   1  'Simple Combo
            TabIndex        =   40
            Top             =   780
            Width           =   312
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Search FA"
            Height          =   210
            Left            =   8340
            TabIndex        =   39
            Top             =   2484
            Visible         =   0   'False
            Width           =   276
         End
         Begin VB.CommandButton cmdShow1 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4740
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   660
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   11400
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1260
            Width           =   1365
         End
         Begin VB.TextBox txtRem 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   948
            Left            =   5928
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   696
            Width           =   5388
         End
         Begin VB.CommandButton cmdupdatep 
            Caption         =   "Party"
            Height          =   195
            Left            =   -61236
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   84
         End
         Begin VB.ComboBox closingcr 
            Height          =   288
            Left            =   12780
            Style           =   1  'Simple Combo
            TabIndex        =   34
            Top             =   1260
            Width           =   312
         End
         Begin VB.ComboBox cboStation 
            Height          =   288
            Left            =   9960
            TabIndex        =   33
            Top             =   2460
            Width           =   2112
         End
         Begin VB.ComboBox COMBOGENLEDGER 
            Height          =   288
            Left            =   -62760
            TabIndex        =   32
            Text            =   "SUNDRY DEBTORS"
            Top             =   4320
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.CommandButton cmdShowClosing 
            BackColor       =   &H00FAEFC9&
            Caption         =   "Closing"
            Height          =   435
            Left            =   -63336
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1176
            Width           =   2010
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -63336
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   732
            Width           =   2010
         End
         Begin VB.TextBox txtamount 
            Height          =   315
            Left            =   -62736
            TabIndex        =   29
            Top             =   3936
            Width           =   1350
         End
         Begin VB.ComboBox cboStation1 
            Height          =   288
            Left            =   -62736
            TabIndex        =   27
            Top             =   3645
            Width           =   1365
         End
         Begin VB.TextBox txtalfa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   12960
            MaxLength       =   8
            TabIndex        =   26
            Top             =   2460
            Width           =   315
         End
         Begin VB.CommandButton cmdMain 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Exit"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   10584
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   2865
            Width           =   876
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   2865
            Width           =   864
         End
         Begin VB.CommandButton cmdRefresh 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Refresh"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2865
            Width           =   888
         End
         Begin VB.CommandButton cmdDel 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Delete"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7116
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2865
            Width           =   888
         End
         Begin VB.CommandButton cmdModify 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Modify"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6264
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   2865
            Width           =   840
         End
         Begin VB.CommandButton cmdprintalf 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Print Alphabet"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5304
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   2865
            Width           =   948
         End
         Begin VB.CheckBox selectAll 
            Caption         =   "Select All"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   12072
            TabIndex        =   17
            Top             =   2460
            Width           =   924
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FAEFC9&
            Caption         =   "&Outstanding List"
            Enabled         =   0   'False
            Height          =   408
            Left            =   -63444
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   6288
            Width           =   2010
         End
         Begin VB.Frame dataTrans 
            Height          =   555
            Left            =   -74040
            TabIndex        =   13
            Top             =   8700
            Width           =   6015
            Begin VB.TextBox txtPath 
               Height          =   315
               Left            =   120
               TabIndex        =   15
               Top             =   180
               Width           =   4875
            End
            Begin VB.CommandButton cmdPath 
               Caption         =   "&Path"
               Height          =   315
               Left            =   5160
               TabIndex        =   14
               Top             =   180
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdAson 
            BackColor       =   &H00FAEFC9&
            Caption         =   "Closing  As On Date"
            Height          =   396
            Left            =   -63408
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2844
            Width           =   2010
         End
         Begin VB.CheckBox Check2 
            Caption         =   "State Wise"
            Height          =   285
            Left            =   -62736
            TabIndex        =   10
            Top             =   3375
            Width           =   1320
         End
         Begin VB.Frame Frame3 
            Caption         =   "Order By"
            Height          =   1290
            Left            =   -63468
            TabIndex        =   6
            Top             =   6828
            Visible         =   0   'False
            Width           =   1995
            Begin VB.OptionButton PartyWise 
               Caption         =   "Party Wise"
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
               Left            =   60
               TabIndex        =   9
               Top             =   360
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton citywise 
               Caption         =   "City Wise"
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
               Left            =   60
               TabIndex        =   8
               Top             =   675
               Width           =   1335
            End
            Begin VB.OptionButton Balance 
               Caption         =   "Balance Wise"
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
               Left            =   60
               TabIndex        =   7
               Top             =   990
               Width           =   1575
            End
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FAEFC9&
            Caption         =   "Mobile No.Export To Notepad"
            Height          =   435
            Left            =   -62715
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   4275
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.CheckBox Check_ClosingDesc 
            Caption         =   "Closing Desc. Order"
            Height          =   255
            Left            =   -63336
            TabIndex        =   4
            Top             =   384
            Width           =   1755
         End
         Begin MSComCtl2.DTPicker dateason 
            Height          =   312
            Left            =   -62676
            TabIndex        =   12
            Top             =   2484
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   550
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   494206979
            CurrentDate     =   39979
         End
         Begin MSComDlg.CommonDialog comdio 
            Left            =   -74640
            Top             =   8820
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Crystal.CrystalReport crpt 
            Left            =   9216
            Top             =   8856
            _ExtentX        =   593
            _ExtentY        =   593
            _Version        =   348160
            PrintFileUseRptNumberFmt=   -1  'True
            PrintFileUseRptDateFmt=   -1  'True
            PrintFileLinesPerPage=   60
            WindowShowCloseBtn=   -1  'True
            WindowShowSearchBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin MSComCtl2.DTPicker RecDates 
            Height          =   288
            Left            =   3396
            TabIndex        =   1
            Top             =   636
            Width           =   1296
            _ExtentX        =   2286
            _ExtentY        =   508
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   494206979
            UpDown          =   -1  'True
            CurrentDate     =   37701
         End
         Begin VSFlex7Ctl.VSFlexGrid vs1 
            Height          =   5292
            Left            =   60
            TabIndex        =   52
            Top             =   3516
            Width           =   14748
            _cx             =   26014
            _cy             =   9334
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16777215
            ForeColor       =   16711680
            BackColorFixed  =   16251308
            ForeColorFixed  =   255
            BackColorSel    =   14155775
            ForeColorSel    =   16744448
            BackColorBkg    =   16251308
            BackColorAlternate=   16777215
            GridColor       =   255
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   3
            GridLines       =   9
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   640
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   -1  'True
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
            Editable        =   0
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
         Begin VSFlex7Ctl.VSFlexGrid vsop 
            Height          =   8160
            Left            =   -74940
            TabIndex        =   53
            Top             =   480
            Width           =   11148
            _cx             =   19664
            _cy             =   14393
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16251308
            ForeColor       =   16711680
            BackColorFixed  =   16251308
            ForeColorFixed  =   255
            BackColorSel    =   16448755
            ForeColorSel    =   16744448
            BackColorBkg    =   16251308
            BackColorAlternate=   16251308
            GridColor       =   255
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   8
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   340
            RowHeightMax    =   0
            ColWidthMin     =   800
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
            ScrollTips      =   -1  'True
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
            ExplorerBar     =   7
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
         Begin MSMask.MaskEdBox date1 
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
            Left            =   -62490
            TabIndex        =   54
            Top             =   4320
            Visible         =   0   'False
            Width           =   120
            _ExtentX        =   212
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox date2 
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
            Left            =   -63024
            TabIndex        =   55
            Top             =   5964
            Width           =   1176
            _ExtentX        =   2074
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   315
            Left            =   -73935
            TabIndex        =   94
            Top             =   870
            Width           =   1335
            _ExtentX        =   2350
            _ExtentY        =   550
            _Version        =   393216
            Format          =   127926273
            CurrentDate     =   38845
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   300
            Left            =   -73950
            TabIndex        =   95
            Top             =   540
            Width           =   1350
            _ExtentX        =   2392
            _ExtentY        =   529
            _Version        =   393216
            Format          =   127926273
            CurrentDate     =   38845
         End
         Begin VSFlex7Ctl.VSFlexGrid vs 
            Height          =   7320
            Left            =   -74910
            TabIndex        =   96
            Top             =   1920
            Width           =   13365
            _cx             =   23574
            _cy             =   12912
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16777209
            ForeColor       =   16711680
            BackColorFixed  =   16777173
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777166
            BackColorAlternate=   16777209
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmBillList.frx":00EF
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
            ExplorerBar     =   7
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
         Begin VSFlex7Ctl.VSFlexGrid vs_document 
            Height          =   2844
            Left            =   14868
            TabIndex        =   150
            Top             =   3528
            Width           =   2292
            _cx             =   4043
            _cy             =   5016
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   16761024
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16761992
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   450
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmBillList.frx":01CF
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
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   -1  'True
            WordWrap        =   -1  'True
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
         Begin VSFlex7Ctl.VSFlexGrid VS_sale 
            Height          =   2016
            Left            =   14868
            TabIndex        =   154
            Top             =   6768
            Width           =   2292
            _cx             =   4043
            _cy             =   3556
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   16761024
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16761992
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   450
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmBillList.frx":0258
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
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   -1  'True
            WordWrap        =   -1  'True
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
         Begin VB.Label lblCAF 
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   492
            Index           =   1
            Left            =   14868
            TabIndex        =   151
            Top             =   2916
            Width           =   2316
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Cheque No"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   228
            Left            =   13356
            TabIndex        =   147
            Top             =   2304
            Width           =   1452
         End
         Begin VB.Label lblfrt 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   10.2
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   5832
            TabIndex        =   145
            Top             =   2376
            Width           =   2172
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Series Wise Discount"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   420
            Index           =   1
            Left            =   9792
            MousePointer    =   1  'Arrow
            TabIndex        =   139
            Top             =   8892
            Visible         =   0   'False
            Width           =   4764
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Series Wise Discount"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   312
            Index           =   2
            Left            =   9756
            TabIndex        =   140
            Top             =   8856
            Visible         =   0   'False
            Width           =   4764
         End
         Begin VB.Label lblTotalRecord 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   135
            TabIndex        =   128
            Top             =   8955
            Width           =   1965
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   2
            Height          =   960
            Left            =   -63636
            Top             =   2412
            Width           =   2580
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   2
            Height          =   1896
            Left            =   -63636
            Top             =   3336
            Width           =   2580
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   2
            Height          =   1740
            Left            =   -63636
            Top             =   660
            Width           =   2580
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   2
            Height          =   1596
            Left            =   -63636
            Top             =   5208
            Width           =   2580
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CASH PARTY"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   420
            Index           =   20
            Left            =   5832
            TabIndex        =   123
            Top             =   1764
            Visible         =   0   'False
            Width           =   2316
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CASH PARTY"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Index           =   21
            Left            =   5805
            TabIndex        =   122
            Top             =   1755
            Visible         =   0   'False
            Width           =   3075
         End
         Begin VB.Label Label22 
            Caption         =   "Party Outstanding List To whom goods are not supply from this date"
            Height          =   552
            Left            =   -63516
            TabIndex        =   121
            Top             =   5316
            Width           =   2460
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "As On:"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   12420
            TabIndex        =   117
            Top             =   3120
            Width           =   720
         End
         Begin VB.Label lblTOD 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   120
            TabIndex        =   115
            Top             =   8580
            Width           =   3675
         End
         Begin VB.Label Label1 
            Caption         =   "From Date "
            Height          =   240
            Index           =   0
            Left            =   -74790
            TabIndex        =   98
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label Label2 
            Caption         =   "To Date"
            Height          =   255
            Left            =   -74790
            TabIndex        =   97
            Top             =   930
            Width           =   1245
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Desc."
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   252
            Left            =   96
            TabIndex        =   76
            Top             =   1440
            Width           =   1080
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Total "
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   5205
            TabIndex        =   75
            Top             =   8955
            Width           =   795
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E98A0A&
            BackStyle       =   0  'Transparent
            Caption         =   "Esc To Exit"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   252
            Left            =   8472
            TabIndex        =   74
            Top             =   384
            Width           =   972
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount "
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   90
            TabIndex        =   73
            Top             =   2355
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Date :"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   252
            Left            =   2760
            TabIndex        =   72
            Top             =   684
            Width           =   756
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Rec. No:"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   90
            TabIndex        =   71
            Top             =   690
            Width           =   975
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "P.Name "
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   252
            Left            =   96
            TabIndex        =   70
            Top             =   1080
            Width           =   1044
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Opening"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   276
            Left            =   11400
            TabIndex        =   69
            Top             =   600
            Width           =   960
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Closing"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   276
            Left            =   11400
            TabIndex        =   68
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "District"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   276
            Left            =   8616
            TabIndex        =   67
            Top             =   2520
            Width           =   1356
         End
         Begin VB.Label abc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Height          =   312
            Left            =   -63336
            TabIndex        =   66
            Top             =   1656
            Width           =   2016
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   276
            Left            =   -63564
            TabIndex        =   65
            Top             =   3996
            Width           =   900
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone "
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   90
            TabIndex        =   64
            Top             =   1845
            Width           =   1335
         End
         Begin VB.Label phone 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1320
            TabIndex        =   63
            Top             =   1875
            Width           =   6975
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Press F1 To Set Description"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   192
            Left            =   1260
            TabIndex        =   62
            Top             =   384
            Width           =   2952
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Press F5 For Up && Down"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   408
            Left            =   3960
            TabIndex        =   61
            Top             =   2340
            Width           =   1416
         End
         Begin VB.Label CrLebel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3480
            TabIndex        =   60
            Top             =   8940
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label drLebel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2100
            TabIndex        =   59
            Top             =   8940
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "As On:"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   276
            Left            =   -63456
            TabIndex        =   58
            Top             =   2592
            Width           =   720
         End
         Begin VB.Label lblStation 
            BackStyle       =   0  'Transparent
            Caption         =   "Station"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   276
            Left            =   -63564
            TabIndex        =   57
            Top             =   3708
            Width           =   1200
         End
         Begin VB.Label lblCR 
            Height          =   315
            Left            =   10320
            TabIndex        =   56
            Top             =   1140
            Visible         =   0   'False
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frmBillList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim rs9 As New ADODB.Recordset
Dim bb As Boolean
Dim bb2 As Boolean
Dim rss As New ADODB.Recordset
Dim from_date As Date
Dim I As Integer
Dim str_CREDITORS As Boolean
Dim col_name As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim search_v As Boolean
Dim ch_ As Boolean
Dim to_date As Date
Dim SessionLastDate As Date
Dim kk As Integer
Dim bb1 As Boolean
Dim str1 As New ADODB.Recordset
Dim rs_don As New ADODB.Recordset
Dim con_don As New ADODB.Connection
Dim rptid As Integer
Dim fdate1_
Dim CON_next_ As New ADODB.Connection
Dim dt_str, dt_strR As String
Dim user_id As String
Dim CASH_profile As String
Dim str_don As String

Dim saleAmt1, RetAmt1 As Double

Dim SessionLastDate1 As Date
Dim rs_closing As ADODB.Recordset
Sub CalculateTotalDrCrNew()

On Error GoTo aa1

Dim Balance As Long
Dim dr1, cr1, prbal
Dim rs_1 As New ADODB.Recordset
Dim Str
lblTotalRecord.Caption = ""
txtOp = 0
txtBalance = 0
Str = ""
dr1 = 0
cr1 = 0
txtClosing.text = 0
txtcr.text = 0
If RS.State = 1 Then RS.close
RS.Open "select Op,drcr,YEAROPENING from SLEDGER where " & stringyear & " and SUBLEDGER='" & cboParty.text & "'", con
If RS.EOF = False Then

If lblCr = "cr" Then
   txtOp.text = Format(RS.Fields(2).value, "0.00")
   cmdSave.Enabled = False
Else
   txtOp.text = Format(RS.Fields(0).value, "0.00")
   cmdSave.Enabled = True
End If

txtOp.text = Abs(txtOp.text)

If Len(RS.Fields("drcr").value) >= 2 Then
If UCase(RS.Fields("drcr").value) = UCase("dr") Then
cboop.text = "Dr"
Else
cboop.text = "Cr"
End If
End If

Else
txtOp.text = 0
End If

'=====================================================
  If vs1.rows <= 1 Then
        txtBalance = txtOp
        closingcr = cboop
     Exit Sub
  End If
'=====================================================


If cboop.text = "Dr" Then
   dr1 = (Val(txtOp.text) + Val(vs1.TextMatrix(1, 4)))
   cr1 = Val(vs1.TextMatrix(1, 5))
Else
   cr1 = (Val(txtOp.text) + Val(vs1.TextMatrix(1, 5)))
   dr1 = Val(vs1.TextMatrix(1, 4))
End If


prbal = dr1 - cr1
If prbal < 0 Then
vs1.TextMatrix(1, 6) = Format(-1 * prbal, "0.00")
vs1.TextMatrix(1, 7) = "Cr"
Else
vs1.TextMatrix(1, 6) = Format(prbal, "0.00")
vs1.TextMatrix(1, 7) = "Dr"
End If


For I = 1 To vs1.rows - 1

If vs1.TextMatrix(I, 0) <> "" Then

txtClosing.text = (Val(txtClosing.text) + Val(vs1.TextMatrix(I, 4)))
txtcr.text = (Val(txtcr.text) + Val(vs1.TextMatrix(I, 5)))
'-----Balance---------------
If I >= 2 Then
dr1 = Val(vs1.TextMatrix(I, 4))
cr1 = (-1 * Val(vs1.TextMatrix(I, 5)))
bal = dr1 + cr1
If Str = "Cr" Then
bal = prbal + bal
Else
bal = prbal + bal
End If
If bal < 0 Then
vs1.TextMatrix(I, 6) = Format(-1 * bal, "0.00")
vs1.TextMatrix(I, 7) = "Cr"
Else
vs1.TextMatrix(I, 6) = Format(bal, "0.00")
vs1.TextMatrix(I, 7) = "Dr"
End If
prbal = bal
Str = vs1.TextMatrix(I, 7)
End If

End If

Next



txtClosing.text = Format(txtClosing.text, "0.00")
sum11 = 0

txtcr.text = Format(txtcr.text, "0.00")
If cboop.text = "Dr" Then
  txtClosing.text = Format((CDbl(txtClosing.text)), "0.00")
Else
  txtcr.text = Format((CDbl(txtcr.text)), "0.00")
End If

txtBalance.text = prbal
closingcr.text = Str



txtBalance.text = Format(txtBalance.text, "0.00")
lblTotalRecord.Caption = "Tot.Rows : " & vs1.rows


If txtClosing.text = 0 And txtcr.text = 0 Then
   vs1.TextMatrix(1, 6) = ""
   vs1.TextMatrix(1, 7) = ""
   lblTotalRecord.Caption = "Tot.Rows : " & 0
End If


Exit Sub
aa1:
MsgBox "" & err.DESCRIPTION, vbCritical

End Sub
Sub genrateCloasing()


Set rs_closing = New ADODB.Recordset
Set rs_closing = con.Execute("exec spPartyClosing")


End Sub
Sub PartyLedgerNew()

On Error GoTo view_

vs1.Clear
setWidth

Dim pname As String

pname = Trim(cboParty.text)

user_id = Trim((Sys_user_ + Str(UId)))


'===============================================================================
'====================================================================================

con.Execute "delete from tmpDonnationnew where uid='" & user_id & "'"
con.Execute "delete from tmpSaladjust where uid='" & user_id & "'"


con.Execute "delete from templedger6 where userid='" & user_id & "'"
con.Execute "exec tmpLedgerNew '" & session & "','" & pname & "','" & user_id & "'"


con.Execute "INSERT INTO templedger6 (Balance,drcr,party,billtype,rptid,rptype,setupid,fyear,district,userid,states,Party1)  SELECT op,drcr,subledger,'Opening',1,'" & cboStation.text & "',setupid,fyear,ADDRESS3,'" & user_id & "',states,DESCFORINVOICE from sledger where substring(subledger,1,5)='" & Mid(pname, 1, 5) & "'  group by op,subledger,drcr,setupid,Fyear,ADDRESS3,states,DESCFORINVOICE  HAVING  op <> 0"
con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname,billtype1,sdiscount,todno,toddate,scid,bookgp)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales Bilty No-' + BILTYNO + ',Bundle-' + bundles ,netamount,BAA,SUBLEDGER,fyear,setupid,'" & user_id & "',district,'1','" & cboStation.text & "',states,Party,'',scname,'I',sdiscount,todid,toddate,scid,gpname  from invoiceaQry where substring(Subledger,1,5)='" & Mid(pname, 1, 5) & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname,billtype1,todno,toddate,scid,bookgp) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item Bilty No-' + BILTYNO + ',Bundle-' + bundles,BAA,netamount,SUBLEDGER,fyear,setupid,'" & user_id & "',district,'1','" & cboStation.text & "',states,Party,'',scname,'C',todid,toddate,scid,gpname from CREDITAQry where substring(Subledger,1,5)='" & Mid(pname, 1, 5) & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"

con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,sdiscount)" & _
"SELECT  CASHA.INVOICEDATE,'C/M',CASHA.INVOICENO,'Cash Memo',CASHA.NETAMOUNT,CASHA.BAA,CASHA.cashpartyname,CASHA.Fyear," & _
"CASHA.setupid,'" & user_id & "',SLEDGER.ADDRESS3,'1','" & cboStation.text & "',SLEDGER.states,SLEDGER.DESCFORINVOICE,'',casha.sdiscount " & _
"FROM CASHA INNER JOIN SLEDGER ON CASHA.SUBLEDGER = SLEDGER.SUBLEDGER where substring(CASHA.SUBLEDGER,1,5)='" & Mid(pname, 1, 5) & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"

'-
con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,Billtype1,todno,toddate)" & _
" SELECT CNF1A.CND,'CN',CNF1A.cnn,'Credit Note ' + desc_ ,0,CNF1A.NA,CNF1A.psld,CNF1A.Fyear," & _
"CNF1A.setupid,'" & user_id & "',SLEDGER.ADDRESS3,'1','" & cboStation.text & "',SLEDGER.states,SLEDGER.DESCFORINVOICE,'','CN',todid,toddate " & _
"FROM  dbo.CNF1A INNER JOIN SLEDGER ON CNF1A.psld = dbo.SLEDGER.SUBLEDGER where  substring(SLEDGER.SUBLEDGER,1,5) ='" & Mid(pname, 1, 5) & "' and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) "


con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,Billtype1)" & _
"SELECT DNFA.DND,'DN',DNFA.Dnn,'Debit Note ' + desc_,DNFA.NA,0,DNFA.psld,DNFA.Fyear," & _
"DNFA.setupid,'" & user_id & "',SLEDGER.ADDRESS3,'1','" & cboStation.text & "',SLEDGER.states,SLEDGER.DESCFORINVOICE,'','DN' " & _
"FROM DNFA INNER JOIN SLEDGER ON DNFA.psld = SLEDGER.SUBLEDGER where substring(SLEDGER.SUBLEDGER,1,5) ='" & Mid(pname, 1, 5) & "' and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)"

'-'

con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) " & _
" SELECT a.Dates,'J',a.RecNo, a.Particullar, a.Dr, a.Cr,a.PartyName,a.fyear,a.setupid,'" & user_id & "'," & _
" b.ADDRESS3,'1','" & cboStation.text & "',b.states,b.DESCFORINVOICE,'' FROM ReceiveIssueParty as a INNER JOIN " & _
" SLEDGER as b ON a.PartyName = b.SUBLEDGER where substring(a.PartyName,1,5) ='" & Mid(pname, 1, 5) & "' and convert(smalldatetime,DATEs,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and a.firm='" & firm & "'  order by dates,recno"


'===================================================================================
'--------Donation Details-----------------------------------------------------------

con.Execute "update a set a.repname = tmpDonnationnew.entryNo,a.adjdate  = tmpDonnationnew.Dates  from tmpDonnationnew " & _
"INNER JOIN templedger6 as a ON (a.bill = tmpDonnationnew.billno and a.Billtype1  = tmpDonnationnew.category and SUBSTRING(a.party,1,5) = SUBSTRING(tmpDonnationnew.party,1,5) and (a.fyear = tmpDonnationnew.fyear))"


'-----End Donation Details-------------------------------------------------------------

con.Execute "update a set a.spno = tmpSaladjust.entryNo,a.spdate  = tmpSaladjust.Dates  from tmpSaladjust " & _
"INNER JOIN templedger6 as a ON (a.bill = tmpSaladjust.billno and a.Billtype1  = tmpSaladjust.category and  SUBSTRING(a.party, 1, 5) = SUBSTRING(tmpSaladjust.party, 1, 5) and a.fyear = tmpSaladjust.fyear)"


con.Execute "update a set a.spno = tmpSaladjust.entryNo,a.spdate  = tmpSaladjust.Dates  from tmpSaladjust " & _
"INNER JOIN templedger6 as a ON (a.bill = tmpSaladjust.billno and  SUBSTRING(a.party, 1, 5) = SUBSTRING(tmpSaladjust.party, 1, 5) and a.fyear = tmpSaladjust.fyear and a.billtype='C/M')"


'-----End Adjustment Details--------------------------------------------------------

con.Execute "exec forUpdateCrNoIn_Ledger"

'-----------------------------------------------------------------------------------

Dim rs_ap As ADODB.Recordset
Set rs_ap = New ADODB.Recordset


'Set rs_ap = con.Execute("exec searchList '" & "findappno" & "'")

'-----------------------------------------------------------------------------------




'-----------------------------------------------------------------------------------
Dim ff As New ADODB.Recordset
Dim CR, dr As Double

CR = 0
dr = 0
s10 = ""
Dim bookgp As String

If ff.State = 1 Then ff.close
ff.Open "select Billtype,bill,dates,des,dr,cr,Balance,drcr,repname,adjDate,SPNO,SPDATE,TODNO,TODDATE,sdiscount,scname,bookgp from templedger6  where (userid='" & user_id & "' and SUBSTRING(party,1,5)='" & Mid(pname, 1, 5) & "' and dates is not null) order by dates,bill", con
vs1.rows = ff.RecordCount + 1
For J = 1 To vs1.rows - 1
 If ff.EOF = False Then
     s10 = ""
     vs1.TextMatrix(J, 0) = ff.Fields(0).value
     vs1.TextMatrix(J, 1) = ff.Fields(1).value
     vs1.TextMatrix(J, 2) = ff.Fields(2).value
     
     If Not IsNull(ff!bookgp) Then
        bookgp = ff!bookgp & ""
     End If
     '----------------------------------------
        If Not IsNull(ff!sdiscount) Then
        If Val(ff!sdiscount) > 0 Then
           s10 = "(DS -" & ff!sdiscount & ")"
        End If
        End If
          
        If Not IsNull(ff!scname) Then
        If Len(ff!scname) > 0 Then
           If s10 = "" Then
              s10 = "SCHOOL-" & ff!scname
           Else
              s10 = s10 & ",SCHOOL-" & ff!scname
           End If
        End If
        End If
     '----------------------------------------
     If s10 <> "" Then
        vs1.TextMatrix(J, 3) = IIf(IsNull(ff.Fields(3).value), "-", ff.Fields(3).value) & ", " & s10
     Else
        vs1.TextMatrix(J, 3) = IIf(IsNull(ff.Fields(3).value), "-", ff.Fields(3).value)
     End If
     
     If (vs1.TextMatrix(J, 0) = "I" Or vs1.TextMatrix(J, 0) = "CI") Then
        If Len(bookgp) > 0 Then
           vs1.TextMatrix(J, 3) = vs1.TextMatrix(J, 3) & " (" & bookgp & ")"
        End If
        
        
     End If
     
     vs1.TextMatrix(J, 4) = Format(ff.Fields(4).value, "0.00")
     vs1.TextMatrix(J, 5) = Format(ff.Fields(5).value, "0.00")
     vs1.TextMatrix(J, 6) = Format(ff.Fields(6).value, "0.00")
     
     
     vs1.TextMatrix(J, 8) = ff.Fields("repname").value & "-" & ff.Fields("adjDate").value
     
     If (vs1.TextMatrix(J, 0) = "I" Or vs1.TextMatrix(J, 0) = "CI") Then
     
        ''' new code --------
        
        Set rs9 = New ADODB.Recordset
        If vs1.TextMatrix(J, 0) = "I" Then
           rs9.Open "select entryNo as SPNO,Dates as SPDate,billno from tmpSaladjust where category='I' and party='" & pname & "' and billno='" & vs1.TextMatrix(J, 1) & "' and fyear='" & session & "' group by entryNo,Dates,billno", con
        ElseIf vs1.TextMatrix(J, 0) = "CI" Then
            rs9.Open "select entryNo as SPNO,Dates as SPDate,billno from tmpSaladjust where category='I' and party='" & pname & "' and billno='" & vs1.TextMatrix(J, 1) & "' and fyear='" & session & "' group by entryNo,Dates,billno", con
        End If
        
        If rs9.RecordCount > 1 Then
           
           While rs9.EOF = False
                If vs1.TextMatrix(J, 9) = "" Then
                  vs1.TextMatrix(J, 9) = rs9.Fields("SPNO").value & "-" & rs9.Fields("SPDate").value
                Else
                  vs1.TextMatrix(J, 9) = vs1.TextMatrix(J, 9) & "    " & rs9.Fields("SPNO").value & "-" & rs9.Fields("SPDate").value
                End If
                rs9.MoveNext
           Wend
           
        Else

           If Not IsNull(ff.Fields("SPDate").value) Then
              vs1.TextMatrix(J, 9) = ff.Fields("SPNO").value & "-" & ff.Fields("SPDate").value
           Else
              vs1.TextMatrix(J, 9) = ff.Fields("SPNO").value & ""
           End If
        
        End If
     
     Else
     
     
        vs1.TextMatrix(J, 8) = ff.Fields("repname").value & "-" & ff.Fields("adjDate").value
        If Not IsNull(ff.Fields("SPDate").value) Then
           vs1.TextMatrix(J, 9) = ff.Fields("SPNO").value & "-" & ff.Fields("SPDate").value
        Else
           vs1.TextMatrix(J, 9) = ff.Fields("SPNO").value & ""
        End If
     
     
     
     End If
     
     
     vs1.TextMatrix(J, 10) = ff.Fields("todno").value & "-" & ff.Fields("todDate").value
     
     CR = CR + ff.Fields(4).value
     dr = dr + ff.Fields(5).value
     
     If vs1.TextMatrix(J, 1) = 4366 Then
     '   MsgBox "s"
     End If
     
     If Mid(session, 6) >= 18 Then
     
            If vs1.TextMatrix(J, 0) = "I" Then
               
               If session = "2018-19" Then
                    If rs_ap.State = 1 Then rs_ap.close
                    rs_ap.Open "SELECT distinct AppNo,INVOICENO FROM ApprovalDet where INVOICENO=" & ff.Fields(1).value & " and fyear='" & session & "'", con
                    If rs_ap.EOF = False Then
                       vs1.TextMatrix(J, 11) = rs_ap.Fields("appno").value
                    End If
               Else
                    If rs_ap.State = 1 Then rs_ap.close
                    rs_ap.Open "SELECT appno FROM invoicea where INVOICENO=" & ff.Fields(1).value & " and fyear='" & session & "' and subledger='" & cboParty.text & "'", con
                    If rs_ap.EOF = False Then
                       vs1.TextMatrix(J, 11) = rs_ap.Fields("appno").value & ""
                    End If
               End If
               
           End If
           
       End If
     
     ff.MoveNext
 End If
Next

'setwidth

txtcr.text = Format(Round(dr, 0), "0.00")
txtClosing.text = Format(Round(CR, 0), "0.00")

Exit Sub
view_:
MsgBox "" & err.DESCRIPTION

End Sub
Sub vsIni()


If Option2_donation.value = True Then
   
   ''If ch_din = "y" Then
     vs.Cols = 9
     vs.FormatString = "SNo|Sp.No|Date|School Name|>Amount|^Authorized||Status|Update Sponsorship"
   ''Else
    '' vs.Cols = 8
    '' vs.FormatString = "SNo|Sponsorship No|Date|School Name|>Amount|^bauthorized"
   ''End If
   
   vs.ColWidth(0) = 600
   vs.ColWidth(1) = 600
   vs.ColWidth(2) = 1100
   vs.ColWidth(3) = 4400
   vs.ColWidth(4) = 1000
   vs.ColWidth(5) = 900
   
   ''If ch_din = "y" Then
      vs.ColWidth(6) = 6
      vs.ColWidth(7) = 1500
      vs.ColComboList(7) = "Full Paid|Half Paid|Advance|Pending| "
   ''Else
   ''   vs.ColWidth(7) = 0
   ''End If
 
   
   Exit Sub
 
 End If
  
     
 If crdit.value = False Then
 
   vs.Cols = 7
   vs.FormatString = "SNo|Bill/Credit Bill No|Date|Party Name|>Amount|^Authorized|Balance"
   vs.ColWidth(0) = 900
   vs.ColWidth(1) = 1500
   vs.ColWidth(2) = 1100
   vs.ColWidth(3) = 5400
   vs.ColWidth(4) = 1400
   vs.ColWidth(5) = 1000
   vs.ColWidth(6) = 1200
 
 
 Else
 
      DoEvents
      vs.Cols = 7
      DoEvents
      vs.FormatString = "SNo|Bill/Credit Bill No|Date|Party Name|>Amount|^Authorized|Balance"
      vs.ColWidth(0) = 900
      vs.ColWidth(1) = 1500
      vs.ColWidth(2) = 1200
      vs.ColWidth(3) = 4200
      vs.ColWidth(4) = 1400
      vs.ColWidth(5) = 1000
      vs.ColWidth(6) = 1200
      DoEvents
      DoEvents
 End If
   
End Sub

Private Sub All_Click()
If All.value = True Then
    Call cmdshow_Click
End If

End Sub

Private Sub autho_Click()
If autho.value = True Then
    Call cmdshow_Click
End If
End Sub



Private Sub cash_Click()
    If cash.value = True Then
       Call cmdshow_Click
    End If
End Sub

Private Sub cboop_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtdes.SetFocus
   End If
End Sub

Private Sub cboStation_Click()



If Option3_dist.value = True Then
    cboPartyList.Visible = True
    
    If RS.State = 1 Then RS.close
    RS.Open "select distinct(SUBLEDGER) from SLEDGER where " & stringyear & " and DISTCODE='" & cboStation.text & "'", con
    cboPartyList.Clear
    While RS.EOF = False
    cboPartyList.AddItem RS(0)
    RS.MoveNext
    Wend
ElseIf Option2_state.value = True Then
    cboPartyList.Visible = False
    If RS.State = 1 Then RS.close
    RS.Open "select distinct(SUBLEDGER) from SLEDGER where " & stringyear & " and states='" & cboStation.text & "'", con
    cboPartyList.Clear
    k_10 = 0
    While RS.EOF = False
    cboPartyList.AddItem RS(0)
    cboPartyList.Selected(k_10) = True
    k_10 = k_10 + 1
    RS.MoveNext
    Wend
    
ElseIf Option4_rep.value = True Then
    
    cboPartyList.Visible = True
    
    
    If RS.State = 1 Then RS.close
    If cboStation.text = "ALL" Then
       RS.Open "SELECT distinct(SUBLEDGER) FROM SLEDGER", con
    Else
    'RS.Open "SELECT distinct(SUBLEDGER) FROM PartyWiseRepName where AgentName='" & cboStation.Text & "'", con
    RS.Open "SELECT distinct(SUBLEDGER) FROM sledger where (RepName1='" & cboStation.text & "' or RepName2='" & cboStation.text & "' or RepName3='" & cboStation.text & "' or RepName3='" & cboStation.text & "' or RepName4='" & cboStation.text & "' or RepName5='" & cboStation.text & "') AND CHARINDEX('SUNDRY', SUBLEDGER)=0", con
    'where (RepName1='ROHIT SOLOMAN BANERJEE' or RepName2='ROHIT SOLOMAN BANERJEE' or RepName3='ROHIT SOLOMAN BANERJEE' or RepName4='ROHIT SOLOMAN BANE' or RepName5='ROHIT SOLOMAN BANERJEE' )
    
    End If
    cboPartyList.Clear
    While RS.EOF = False
    cboPartyList.AddItem RS(0)
    RS.MoveNext
    Wend

ElseIf Option2_mzn.value = True Then
    
    cboPartyList.Visible = True
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT SUBLEDGER  FROM Managerqry where len(manager)>0 and Manager='" & cboStation.text & "' AND CHARINDEX('SUNDRY', SUBLEDGER)=0 group by SUBLEDGER", con
   
    cboPartyList.Clear
    While RS.EOF = False
      cboPartyList.AddItem RS(0)
      RS.MoveNext
    Wend
End If


End Sub

Private Sub Check1_Click()
    If Check1.value = 1 Then
       'cmdSave.Enabled = False
       cmdDel.Enabled = False
       cmdModify.Enabled = False
    Else
       cmdSave.Enabled = True
       cmdDel.Enabled = True
       cmdModify.Enabled = True
    End If
End Sub

Private Sub Check2_Click()

Dim rs_1 As New ADODB.Recordset

cboStation.Clear
cboStation1.Clear

If Check2.value = 1 Then
    
    lblStation.Caption = "State :"
    
    If rs_1.State = 1 Then rs_1.close
    rs_1.Open "select distinct(states) from SLEDGER where " & stringyear & " and states<>''", con
    While rs_1.EOF = False
    cboStation.AddItem rs_1.Fields(0).value
    cboStation1.AddItem rs_1.Fields(0).value
    rs_1.MoveNext
    Wend

ElseIf Check2.value = 0 Then
    
    lblStation.Caption = "Station :"

    If rs_1.State = 1 Then rs_1.close
    rs_1.Open "select distinct(DISTCODE) from SLEDGER where " & stringyear & " and DISTCODE<>''", con
    While rs_1.EOF = False
    cboStation.AddItem rs_1.Fields(0).value
    cboStation1.AddItem rs_1.Fields(0).value
    rs_1.MoveNext
    Wend


End If

End Sub

Private Sub Check3_filter_Click()


If (Check3_filter.value = 1) Then
   cmdPrint1.Enabled = False
Else
   cmdPrint1.Enabled = True
End If
   
End Sub

Private Sub cmdAson_Click()
showDataAsOn dateAson
End Sub

Private Sub cmdBilty_Click()
frmBiltyReg.Show
End Sub
Private Sub cmdDocument_Click()

Screen.MousePointer = vbHourglass

Dim strProgramName As String
Dim strArgument As String
Dim DocDB As String
Dim con_doc As New ADODB.Connection
Dim caf_ As String

Set con_doc = New ADODB.Connection
DocDB = "Database=ChitraData_2223"
 

 If LCase(server_) = "server" Then
    con_doc.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & DocDB & "; UID=" & sql_user & "; PWD=" & sql_pass
 End If
 
 DoEvents
 DoEvents
 
 
 con_doc.CursorLocation = adUseClient
 If con_doc.State = 1 Then con_doc.close
 con_doc.Open

 DoEvents
 DoEvents




If cboParty.text <> "" Then

Code = Trim(Mid(cboParty.text, 1, 6))
'con.Execute "update setup1 set court='" & Code & "',Mobile_SMS='" & UserName & "'"



con_doc.Execute "update DocumentLink_Code set code='" & Code & "' where UserName='" & UserName & "'"


If RS.State = 1 Then RS.close
RS.Open "select top 1 code from PartyDocument where code='" & Code & "'", con_doc, adOpenDynamic, adLockReadOnly
If RS.EOF = True Then

con_doc.Execute "insert into PartyDocument(code,LinkName) values('" & Code & "','CAF')"
con_doc.Execute "insert into PartyDocument(code,LinkName) values('" & Code & "','GST/PAN')"
con_doc.Execute "insert into PartyDocument(code,LinkName) values('" & Code & "','SECURITY CHEQUES-1')"
con_doc.Execute "insert into PartyDocument(code,LinkName) values('" & Code & "','SECURITY CHEQUES-2')"
con_doc.Execute "insert into PartyDocument(code,LinkName) values('" & Code & "','ADHAR CARD')"
'con_doc.Execute "insert into PartyDocument(code,LinkName) values('" & code & "','MOU')"

End If


dk = Mid(session, 6)
If Val(dk) >= 24 Then
    caf_ = "MOU-" & session
    If RS.State = 1 Then RS.close
    RS.Open "select top 1 code from PartyDocument where code='" & Code & "' and LinkName='" & caf_ & "'", con_doc
    If RS.EOF = True Then
      con_doc.Execute "insert into PartyDocument(code,LinkName) values('" & Code & "','" & caf_ & "')"
    End If

End If



End If

DoEvents
DoEvents

If Code <> "" Then

'If (session = "2022-23") Then

strProgramName = "\\192.168.0.140\blueprintSales\PartyDocument\bin\Debug\MailSystem.exe"
''strProgramName = "C:\SoftwareCode_DotNet_2021\DotNet_soft\SimpleScanningApp\SimpleScanningApp\bin\Debug\SimpleScanningApp.exe"

''strArgument = code & ":" & UserName
strArgument = Code & ":" & UserName & ":" & session


'End If


Call Shell("""" & strProgramName & """ """ & strArgument & """", vbNormalFocus)

'' Call Shell(strProgramName, vbNormalFocus)

End If



Screen.MousePointer = vbDefault

End Sub

Private Sub cmdexit_1_Click()
frmOrderList.Visible = False
End Sub

Private Sub cmdMail_1_Click()

On Error GoTo aa10

Screen.MousePointer = vbHourglass
Dim op, drcr
Dim rs1 As New ADODB.Recordset
Dim rs1_rpt As New ADODB.Recordset
Dim rss_ds As New ADODB.Recordset
Dim inv_str As String

'login.DSN


con.Execute "delete from templedger1 where userid='" & UId & "'"

If rss_ds.State = 1 Then rss_ds.close
rss_ds.Open "select amount,invoiceno from INVOICEC where TEXT='SCHEME DISCOUNT' and AMOUNT>0", con


If rs1_rpt.State = 1 Then rs1_rpt.close
rs1_rpt.Open "select max(rptid) from tempLedger1", con, adOpenDynamic, adLockOptimistic
If IsNull(rs1_rpt(0)) Then
rptid = 9999
Else
rptid = rs1_rpt(0) + 1
End If

If RS.State = 1 Then RS.close
RS.Open "select subledger from SLEDGER where " & stringyear & " and subledger = '" + Trim(cboParty.text) + "'", con
While RS.EOF = False

'==Code For Opening=============================================

If lblCr = "dr" Then

    con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype,setupid,fyear,rptid,UserId)  SELECT op,drcr,subledger,'Opening','" & setupid & "','" & session & "'," & rptid & "," & UId & " from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'   group by op,subledger,drcr HAVING  op <> 0;"
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT op,drcr from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'", con
    If Not IsNull(rs1.Fields(0).value) Then
       op = Val(rs1.Fields(0).value)
       drcr = rs1.Fields(1).value
    Else
       op = 0
    End If

Else

 
    
    con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype,Bill,fyear,setupid,userid)  values(" & Val(txtOp) & ",'" & cboop & "','" & RS.Fields(0).value & "','Opening',0,'" & session & "','" & setupid & "','" & userid & "')"
    op = Val(txtOp)
    drcr = cboop
    
End If


'==============================================


con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER,Fyear,setupid," & rptid & ",'" & UId & "' from INVOICEA where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',BAA,netamount,SUBLEDGER,Fyear,setupid," & rptid & ",'" & UId & "' from CREDITA where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER,Fyear,setupid," & rptid & ",'" & UId & "' from CASHA where  " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD,Fyear,setupid," & rptid & ",'" & UId & "' from dnfa where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party,fyear,setupid,rptid,userid) Select cnd,'CN',cnn,'Credit Note',na,'0',psld,Fyear,setupid," & rptid & ",'" & UId & "' from Cnf1a where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select dates,'J',Recno,Particullar,Dr,CR,PartyName,Fyear,setupid," & rptid & ",'" & UId & "' from ReceiveIssueParty where " & stringyear & " and PartyName='" & RS.Fields(0).value & "' and firm='" & firm & "' order by dates,recno"

If lblCr = "cr" Then

con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select VoucherDate,VoucherType,VoucherNumber,DESCRIPTION,Amount,0,SubLedger,Fyear,setupid," & rptid & ",'" & UId & "' from vouchers where (" & stringyear & " and SubLedger='" & RS.Fields(0).value & "' and DebitorCredit='D') order by VoucherDate,VoucherNumber"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select VoucherDate,VoucherType,VoucherNumber,DESCRIPTION,0,Amount,SubLedger,Fyear,setupid," & rptid & ",'" & UId & "' from vouchers where (" & stringyear & " and SubLedger='" & RS.Fields(0).value & "' and DebitorCredit='C') order by VoucherDate,VoucherNumber"

End If

'===============================================================
If op <> 0 Then
con.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "',UserId='" & UId & "' where " & stringyear & " and party = '" + RS.Fields(0).value + "' and Billtype<>'Opening'"
End If
'===============================================================
Sleep (200)
RS.MoveNext
Wend

'convert(smalldatetime,dates,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)
Dim email As String

If rs1.State = 1 Then rs1.close
rs1.Open "select gledger,Email from SLEDGER where subledger='" & cboParty.text & "'", con
If rs1.EOF = False Then
If rs1!gledger = "SUNDRY DEBTORS" Then
    For J = 1 To vs1.rows - 1
    If vs1.TextMatrix(J, 2) <> "" Then
    
       email = rs1!email & ""
       s_1 = InStr(vs1.TextMatrix(J, 3), "(")
       s_2 = InStr(vs1.TextMatrix(J, 3), ")")
       con.Execute "Update templedger1 set des='" & vs1.TextMatrix(J, 3) & "' where (bill='" + vs1.TextMatrix(J, 1) + "' and Billtype='" + vs1.TextMatrix(J, 0) + "')"
       
       'End If
    End If
    Next
End If
End If

Sleep (10)


''DSNNew
''
''
''
''
''
''    CommonDialog1.ShowPrinter
''
''    crpt.Reset
''    crpt.ReportFileName = App.Path & "\reports\PartyLedger.rpt"
''
''    crpt.ReplaceSelectionFormula "{tempLedgerRpt.UserId}='" & UId & "'"
''
''
''
''    crpt.Connect = constr
''    crpt.WindowShowPrintSetupBtn = True
''    crpt.WindowShowPrintBtn = True
''
''    crpt.Destination = crptToPrinter
''    crpt.Action = 0
''    Screen.MousePointer = vbDefault

   Screen.MousePointer = vbDefault
   PopUpValue6 = cboParty.text
   popupvalue5 = rptid
   popupvalue4 = "PartyLedger.rpt"
   PopUpValue3 = email
   
   frmSendMail.Show 1






Exit Sub
aa10:
Screen.MousePointer = vbDefault
'MsgBox err.DESCRIPTION



End Sub

Private Sub cmdNewPrint_Click()


Screen.MousePointer = vbHourglass

Dim pname As String

con.Execute "delete from templedger5 where userid='" & UId & "'"


If cboStation.text = "ALL" Then


   con.Execute "INSERT INTO templedger5 (Balance,drcr,party,billtype,rptid,rptype,setupid,fyear,district,userid,states,Party1)  SELECT op,drcr,subledger,'Opening',1,'" & cboStation.text & "',setupid,fyear,ADDRESS3," & UId & ",states,DESCFORINVOICE from sledger group by op,subledger,drcr,setupid,Fyear,ADDRESS3,states,DESCFORINVOICE  HAVING  op <> 0"
   con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales Bilty No-' + BILTYNO + ',Bundle-' + bundles ,netamount,BAA,SUBLEDGER,fyear,setupid," & UId & ",district,'1','" & cboStation.text & "',states,Party,AgentName,scname  from invoiceaQry where  convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
   con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',BAA,netamount,SUBLEDGER,fyear,setupid," & UId & ",district,'1','" & cboStation.text & "',states,Party,AgentName,scname from CREDITAQry where convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
    
    con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName)" & _
    "SELECT  CASHA.INVOICEDATE,'C/M',CASHA.INVOICENO,'Cash Memo',CASHA.NETAMOUNT,CASHA.BAA,CASHA.cashpartyname,CASHA.Fyear," & _
    "CASHA.setupid," & UId & ",SLEDGER.ADDRESS3,'1','" & cboStation.text & "',SLEDGER.states,SLEDGER.DESCFORINVOICE,CASHA.AgentName " & _
    "FROM CASHA INNER JOIN SLEDGER ON CASHA.SUBLEDGER = SLEDGER.SUBLEDGER where convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
    
       
    con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName)" & _
    "SELECT CNF1A.CND,'CN',CNF1A.cnn,'Credit Note ' + desc_ ,0,CNF1A.NA,CNF1A.psld,CNF1A.Fyear," & _
    "CNF1A.setupid," & UId & ",SLEDGER.ADDRESS3,'1','" & cboStation.text & "',SLEDGER.states,SLEDGER.DESCFORINVOICE,CNF1A.AgentName " & _
    "FROM  dbo.CNF1A INNER JOIN SLEDGER ON CNF1A.psld = dbo.SLEDGER.SUBLEDGER where convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)"
    
   
    con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName)" & _
    "SELECT DNFA.DND,'CN',DNFA.Dnn,'Debit Note',DNFA.NA,0,DNFA.psld,DNFA.Fyear," & _
    "DNFA.setupid," & UId & ",SLEDGER.ADDRESS3,'1','" & cboStation.text & "',SLEDGER.states,SLEDGER.DESCFORINVOICE,DNFA.AgentName " & _
    "FROM DNFA INNER JOIN SLEDGER ON DNFA.psld = SLEDGER.SUBLEDGER where convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)"
    
    
    
    con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) " & _
    " SELECT a.Dates,'J',a.RecNo, a.Particullar, a.Dr, a.Cr,a.PartyName,a.fyear,a.setupid," & UId & "," & _
    " b.ADDRESS3,'1','" & cboStation.text & "',b.states,b.DESCFORINVOICE,b.repname1 FROM ReceiveIssueParty as a INNER JOIN " & _
    " SLEDGER as b ON a.PartyName = b.SUBLEDGER where convert(smalldatetime,DATEs,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and a.firm='" & firm & "'  order by dates,recno"
    
Else


    For I = 0 To cboPartyList.ListCount - 1
    If cboPartyList.Selected(I) = True Then
        
        pname = cboPartyList.List(I)
        
        con.Execute "INSERT INTO templedger5 (Balance,drcr,party,billtype,rptid,rptype,setupid,fyear,district,userid,states,Party1)  SELECT op,drcr,subledger,'Opening',1,'" & cboStation.text & "',setupid,fyear,ADDRESS3," & UId & ",states,DESCFORINVOICE from sledger where subledger='" & pname & "'  group by op,subledger,drcr,setupid,Fyear,ADDRESS3,states,DESCFORINVOICE  HAVING  op <> 0"
        con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales Bilty No-' + BILTYNO + ',Bundle-' + bundles ,netamount,BAA,SUBLEDGER,fyear,setupid," & UId & ",district,'1','" & cboStation.text & "',states,Party,AgentName,scname  from invoiceaQry where Subledger='" & pname & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
        con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',BAA,netamount,SUBLEDGER,fyear,setupid," & UId & ",district,'1','" & cboStation.text & "',states,Party,AgentName,scname from CREDITAQry where Subledger='" & pname & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
        con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName)" & _
        "SELECT  CASHA.INVOICEDATE,'C/M',CASHA.INVOICENO,'Cash Memo',CASHA.NETAMOUNT,CASHA.BAA,CASHA.cashpartyname,CASHA.Fyear," & _
        "CASHA.setupid," & UId & ",SLEDGER.ADDRESS3,'1','" & cboStation.text & "',SLEDGER.states,SLEDGER.DESCFORINVOICE,CASHA.AgentName " & _
        "FROM CASHA INNER JOIN SLEDGER ON CASHA.SUBLEDGER = SLEDGER.SUBLEDGER where CASHA.SUBLEDGER='" & pname & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
        
   
        
        con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName)" & _
        "SELECT CNF1A.CND,'CN',CNF1A.cnn,'Credit Note ' + desc_ ,0,CNF1A.NA,CNF1A.psld,CNF1A.Fyear," & _
        "CNF1A.setupid," & UId & ",SLEDGER.ADDRESS3,'1','" & cboStation.text & "',SLEDGER.states,SLEDGER.DESCFORINVOICE,CNF1A.AgentName " & _
        "FROM  dbo.CNF1A INNER JOIN SLEDGER ON CNF1A.psld = dbo.SLEDGER.SUBLEDGER where  SLEDGER.SUBLEDGER ='" & pname & "' and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)"
        
          
        con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName)" & _
        "SELECT DNFA.DND,'CN',DNFA.Dnn,'Debit Note',DNFA.NA,0,DNFA.psld,DNFA.Fyear," & _
        "DNFA.setupid," & UId & ",SLEDGER.ADDRESS3,'1','" & cboStation.text & "',SLEDGER.states,SLEDGER.DESCFORINVOICE,DNFA.AgentName " & _
        "FROM DNFA INNER JOIN SLEDGER ON DNFA.psld = SLEDGER.SUBLEDGER where SLEDGER.SUBLEDGER ='" & pname & "' and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)"
        
        
        
        con.Execute "INSERT INTO templedger5 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) " & _
        " SELECT a.Dates,'J',a.RecNo, a.Particullar, a.Dr, a.Cr,a.PartyName,a.fyear,a.setupid," & UId & "," & _
        " b.ADDRESS3,'1','" & cboStation.text & "',b.states,b.DESCFORINVOICE,b.repname1 FROM ReceiveIssueParty as a INNER JOIN " & _
        " SLEDGER as b ON a.PartyName = b.SUBLEDGER where a.PartyName ='" & pname & "' and convert(smalldatetime,DATEs,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and a.firm='" & firm & "'  order by dates,recno"
        
End If
 
Next


End If




con.Execute "UPDATE a SET a.balance = b.op  FROM tempLedger5 AS a " & _
"INNER JOIN SLEDGER AS b ON (a.party = b.subledger)"
con.Execute "UPDATE a SET a.drcr = b.drcr  FROM tempLedger5 AS a " & _
" INNER JOIN SLEDGER AS b ON (a.party = b.subledger)"
 
If Check3_AmountTobecollect.value = 1 Then
  con.Execute "Update templedger5 set AspectedAmt='1'"
End If

 
 
Screen.MousePointer = vbDefault


If MsgBox("Want to Mail ?", vbQuestion + vbYesNo) = vbNo Then

    crpt.Reset
    If cboStation.text = "ALL" Then
       crpt.ReportFileName = rptPath & "\PartyLedgerRepAllNew.rpt"
       crpt.Connect = constr
       crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UId & "'"
    Else
      
      crpt.Connect = constr
      crpt.ReportFileName = rptPath & "\PartyLedgerRepNew.rpt"
      If Option3_dist.value = True Then
       crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UId & "' and {templedgerrpt.district}='" & cboStation.text & "'"
      ElseIf Option2_state.value = True Then
       crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UId & "' and {templedgerrpt.states}='" & cboStation.text & "'"
      ElseIf Option4_rep.value = True Then
       'crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UId & "' and {templedgerrpt.repname}='" & cboStation.Text & "'"
       crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UId & "'"
       'Exit Sub
       ElseIf Option2_mzn.value = True Then
        crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UId & "' and {templedgerrpt.rptype}='" & cboStation.text & "'"
       
      End If
      
    End If
    
    
    crpt.WindowShowPrintSetupBtn = True
    crpt.Formulas(0) = "partyname='" & cboStation.text & "'"
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    crpt.WindowShowRefreshBtn = True
    crpt.WindowShowSearchBtn = True
    crpt.Action = 1

Else

frmSendMail.Show 1


End If



End Sub

Private Sub cmdOrderList_Click()

Screen.MousePointer = vbHourglass
VsOrderList.Clear

Dim s11 As String
s11 = ""
frmOrderList.Visible = True

con.Execute "delete from TmpBook where Login='" & UserName & "'"
If RS.State = 1 Then RS.close
RS.Open "select INVOICENO,BOOKCODE,QUANTITY,BookName,INVOICEDATE,subledger,ScName from OrderBookList where substring([partyname],1,5)='" & Mid(cboParty.text, 1, 5) & "'", con, adOpenDynamic, adLockOptimistic
For I = 1 To RS.RecordCount

If rs1.State = 1 Then rs1.close
rs1.Open "select sum(QUANTITY) from invoiceBQry where (OrderNo=" & RS!invoiceNo & " and bookcode='" & RS!Bookcode & "')", con
If Not IsNull(rs1(0)) Then
   If RS(2) = rs1(0) Then
      s11 = "y"
   Else
      s11 = "n"
   End If
   ii = rs1(0)
Else
   ii = 0
   s11 = "n"
End If

   
con.Execute "insert into TmpBook(BCode,BName,Qty,issueQty,Login,orderno,head,ason,area,repname) values('" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!QUANTITY & "','" & ii & "','" & UserName & "'," & RS!invoiceNo & ",'" & s11 & "','" & RS!invoiceDate & "','" & RS!subledger & "','" & RS!scname & "')"



RS.MoveNext
Next

DoEvents
DoEvents
DoEvents

VsOrderList.Cols = 4

If RS.State = 1 Then RS.close
RS.Open "SELECT OrderNo,Ason,area,sum(Qty),sum(convert(int,issueQty)) as BillQty,repname FROM TmpBook where Login='" & UserName & "' group by OrderNo,Ason,area,repname ", con
For I = 1 To RS.RecordCount
   VsOrderList.TextMatrix(I, 0) = RS!orderNo
   VsOrderList.TextMatrix(I, 1) = RS!Ason
   VsOrderList.TextMatrix(I, 2) = RS!Area
   If RS(3) = RS(4) Then
        For k1 = 0 To 3
          VsOrderList.Cell(flexcpBackColor, I, k1) = vbGreen
        DoEvents
        Next
   End If
   
   VsOrderList.TextMatrix(I, 3) = RS!RepName
   
   DoEvents
   DoEvents
   DoEvents
   RS.MoveNext
Next

VsOrderList.FormatString = "OrderNo|Date|Party Name|School Name"
VsOrderList.ColWidth(0) = 1000
VsOrderList.ColWidth(1) = 1500
VsOrderList.ColWidth(2) = 4000
VsOrderList.ColWidth(3) = 4800


Screen.MousePointer = vbDefault

  
End Sub

Private Sub cmdPath_Click()
Me.comdio.ShowOpen
Me.txtpath.text = Me.comdio.filename
End Sub
Private Sub ExportReportToPDF_(ReportObject As CRAXDRT.Report, ByVal filename As String, ByVal ReportTitle As String)
    
    Dim objExportOptions As CRAXDRT.ExportOptions
 
    ReportObject.ReportTitle = ReportTitle
    
    With ReportObject
        .EnableParameterPrompting = False
        .MorePrintEngineErrorMessages = True
    End With
    
    Set objExportOptions = ReportObject.ExportOptions
    
    With objExportOptions
        .DestinationType = crEDTDiskFile
        .DiskFileName = filename
        .FormatType = crEFTPortableDocFormat
        .PDFExportAllPages = True
    End With
 
    ReportObject.export False
 
End Sub
Public Sub ExportReportToPDF(ReportObject As CRAXDRT.Report, ByVal filename As String, ByVal ReportTitle As String)

    Dim FormatDLLName As String

    ReportObject.ReportTitle = ReportTitle

    With ReportObject
        .EnableParameterPrompting = False
        .MorePrintEngineErrorMessages = True
    End With

    With ReportObject.ExportOptions
        .DestinationType = crEDTDiskFile
        .DiskFileName = filename
        '.FormatType = crEFTExcel80Tabular
        '.FormatType = crEFTCommaSeparatedValues
        '.FormatType = crEFTExcel80
        '.FormatType = crEFTHTML32Standard
        '.FormatType = crEFTHTML40
        .FormatType = crEFTPortableDocFormat
        '.FormatType = crEFTRichText
        '.FormatType = crEFTText
        '.FormatType = crEFTWordForWindows
    End With

 ReportObject.export False

End Sub
Sub pdf()

Dim objCrystal As CRAXDRT.Application
Dim objReport As CRAXDRT.Report

Set objCrystal = New CRAXDRT.Application
ReportFileName = "c:\Report1.rpt"
Set objReport = objCrystal.OpenReport(ReportFileName, 1)
'...code to set report parameters, login information etc...
ExportReportToPDF objReport, "C:\Beds.pdf", "Beds Held"




End Sub
Private Sub cmdPrint_Click()

'pdf

On Error GoTo aa10

Screen.MousePointer = vbHourglass
Dim op, drcr
Dim rs1 As New ADODB.Recordset
Dim rs1_rpt As New ADODB.Recordset
Dim rss_ds As New ADODB.Recordset
Dim inv_str As String

'login.DSN


con.Execute "delete from templedger1 where userid='" & UId & "'"

If rss_ds.State = 1 Then rss_ds.close
rss_ds.Open "select amount,invoiceno from INVOICEC where TEXT='SCHEME DISCOUNT' and AMOUNT>0", con


If rs1_rpt.State = 1 Then rs1_rpt.close
rs1_rpt.Open "select max(rptid) from tempLedger1", con, adOpenDynamic, adLockOptimistic
If IsNull(rs1_rpt(0)) Then
rptid = 9999
Else
rptid = rs1_rpt(0) + 1
End If

If RS.State = 1 Then RS.close
RS.Open "select subledger from SLEDGER where " & stringyear & " and subledger = '" + Trim(cboParty.text) + "'", con
While RS.EOF = False

'==Code For Opening=============================================

If lblCr = "dr" Then

    con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype,setupid,fyear,rptid,UserId)  SELECT op,drcr,subledger,'Opening','" & setupid & "','" & session & "'," & rptid & "," & UId & " from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'   group by op,subledger,drcr HAVING  op <> 0;"
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT op,drcr from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'", con
    If Not IsNull(rs1.Fields(0).value) Then
       op = Val(rs1.Fields(0).value)
       drcr = rs1.Fields(1).value
    Else
       op = 0
    End If

Else

 
    
    con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype,Bill,fyear,setupid,userid)  values(" & Val(txtOp) & ",'" & cboop & "','" & RS.Fields(0).value & "','Opening',0,'" & session & "','" & setupid & "','" & userid & "')"
    op = Val(txtOp)
    drcr = cboop
    
End If


'==============================================


con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER,Fyear,setupid," & rptid & ",'" & UId & "' from INVOICEA where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',BAA,netamount,SUBLEDGER,Fyear,setupid," & rptid & ",'" & UId & "' from CREDITA where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER,Fyear,setupid," & rptid & ",'" & UId & "' from CASHA where  " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD,Fyear,setupid," & rptid & ",'" & UId & "' from dnfa where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party,fyear,setupid,rptid,userid) Select cnd,'CN',cnn,'Credit Note',na,'0',psld,Fyear,setupid," & rptid & ",'" & UId & "' from Cnf1a where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select dates,'J',Recno,Particullar,Dr,CR,PartyName,Fyear,setupid," & rptid & ",'" & UId & "' from ReceiveIssueParty where " & stringyear & " and PartyName='" & RS.Fields(0).value & "' and firm='" & firm & "' order by dates,recno"

If lblCr = "cr" Then

con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select VoucherDate,VoucherType,VoucherNumber,DESCRIPTION,Amount,0,SubLedger,Fyear,setupid," & rptid & ",'" & UId & "' from vouchers where (" & stringyear & " and SubLedger='" & RS.Fields(0).value & "' and DebitorCredit='D') order by VoucherDate,VoucherNumber"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select VoucherDate,VoucherType,VoucherNumber,DESCRIPTION,0,Amount,SubLedger,Fyear,setupid," & rptid & ",'" & UId & "' from vouchers where (" & stringyear & " and SubLedger='" & RS.Fields(0).value & "' and DebitorCredit='C') order by VoucherDate,VoucherNumber"

End If

'===============================================================
If op <> 0 Then
con.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "',UserId='" & UId & "' where " & stringyear & " and party = '" + RS.Fields(0).value + "' and Billtype<>'Opening'"
End If
'===============================================================
Sleep (200)
RS.MoveNext
Wend

'convert(smalldatetime,dates,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)

If rs1.State = 1 Then rs1.close
rs1.Open "select gledger from SLEDGER where subledger='" & cboParty.text & "'", con
If rs1.EOF = False Then
If rs1!gledger = "SUNDRY DEBTORS" Then
    For J = 1 To vs1.rows - 1
    If vs1.TextMatrix(J, 2) <> "" Then
    
    
       s_1 = InStr(vs1.TextMatrix(J, 3), "(")
       s_2 = InStr(vs1.TextMatrix(J, 3), ")")
       con.Execute "Update templedger1 set des='" & vs1.TextMatrix(J, 3) & "' where (bill='" + vs1.TextMatrix(J, 1) + "' and Billtype='" + vs1.TextMatrix(J, 0) + "')"
       
       'End If
    End If
    Next
End If
End If

Sleep (300)


DSNNew




    'CommonDialog1.Flags = 64
    CommonDialog1.ShowPrinter
 
    crpt.Reset
    crpt.ReportFileName = App.Path & "\reports\PartyLedger.rpt"
    
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.UserId}='" & UId & "'"
    
    
   
    crpt.Connect = constr
    crpt.WindowShowPrintSetupBtn = True
    crpt.WindowShowPrintBtn = True
   
    crpt.Destination = crptToPrinter
    crpt.Action = 0
    Screen.MousePointer = vbDefault

'Else
'
'   Screen.MousePointer = vbDefault
'   PopUpValue6 = cboParty.Text
'   popupvalue5 = rptid
'   popupvalue4 = "PartyLedger.rpt"
'   frmSendMail.Show 1
'
'End If




Exit Sub
aa10:
Screen.MousePointer = vbDefault
'MsgBox err.DESCRIPTION


End Sub

Private Sub cmdPrint_Pro_Click()
      
Dim datecondition As String
Dim donationDate As String
      

DoEvents

datecondition = "(convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.value & "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.value & "',103)"
donationDate = "(convert(smalldatetime,DDate,103)>=convert(smalldatetime,'" & fromdate.value & "',103) and convert(smalldatetime,DDate,103)<=convert(smalldatetime,'" & todate.value & "',103)"
      
vs_promotion.ColComboList(0) = "Print|Not"
      
vs_promotion.Clear
vs_promotion.Visible = True

If RS.State = 1 Then RS.close
If All.value = True Then
   RS.Open "select DNo,DDate,ScName,(NetBalance+AdvAmt) as NETAMOUNT,bauthorized,PaymentMode,RepName from DonnationMain where " & donationDate & ") ORDER BY DNo", con
ElseIf autho.value = True Then
   RS.Open "select DNo,DDate,ScName,(NetBalance+AdvAmt) as NETAMOUNT,bauthorized,PaymentMode,RepName from DonnationMain where " & donationDate & ") and bAuthorized=1 ORDER BY DNo", con
Else
   RS.Open "select DNo,DDate,ScName,(NetBalance+AdvAmt) as NETAMOUNT,bauthorized,PaymentMode,RepName from DonnationMain where " & donationDate & ") and bAuthorized=0 ORDER BY DNo", con
End If


vs_promotion.Cols = 9

If RS.EOF = False Then
vs_promotion.rows = RS.RecordCount + 1
For I = 1 To vs_promotion.rows - 1
   DoEvents
   
   vs_promotion.TextMatrix(I, 0) = "Not"
   vs_promotion.TextMatrix(I, 1) = I
   vs_promotion.TextMatrix(I, 2) = RS.Fields(0).value
   vs_promotion.TextMatrix(I, 3) = RS.Fields(1).value
   
   'vs_promotion.TextMatrix(I, 4) = RS.Fields(2).value
   
   If InStr(RS.Fields(2).value, ",") = 0 Then
   Else
      vs_promotion.TextMatrix(I, 4) = Mid(RS.Fields(2).value, 1, InStr(RS.Fields(2).value, ",") - 1)
   End If
   
   
   If InStr(RS.Fields(2).value, ",") = 0 Then
   Else
      vs_promotion.TextMatrix(I, 5) = Mid(RS.Fields(2).value, InStr(RS.Fields(2).value, ",") + 1)
   End If
   
    vs_promotion.TextMatrix(I, 6) = RS.Fields(3).value & ""
    
    vs_promotion.TextMatrix(I, 7) = RS.Fields("PaymentMode").value & ""
    vs_promotion.TextMatrix(I, 8) = RS.Fields("RepName").value & ""
    
    
    
   RS.MoveNext
   DoEvents
Next
Else
   vs_promotion.Clear
   vs_promotion.rows = 2
End If

vs_promotion.FormatString = "Print|SN.|Sp.No|Dates|ScName|Station|Amount|PaymentMode|Representative|Remarks"

vs_promotion.ColWidth(0) = 700
vs_promotion.ColWidth(1) = 500
vs_promotion.ColWidth(2) = 600
vs_promotion.ColWidth(3) = 1200
vs_promotion.ColWidth(4) = 3800
vs_promotion.ColWidth(5) = 1500
vs_promotion.ColWidth(6) = 1000
vs_promotion.ColWidth(7) = 1200
vs_promotion.ColWidth(8) = 1500
vs_promotion.ColWidth(9) = 1500




End Sub

Private Sub cmdPrint1_Click()

crpt.Reset

If Check_ClosingDesc.value = 1 Then
   crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseClosing_descClosing.rpt"
Else
   crpt.ReportFileName = rptPath & "\PartyWiseClosing.rpt"
End If

'crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"

''======================================================================
''======================================================================

If Check2.value = 0 Then

    If cboStation1.text <> "" And txtAmount.text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.text <> "" And txtAmount.text <> "" Then
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.text & ""
    End If
    
    
    ElseIf cboStation1.text = "" And txtAmount.text <> "" Then
    
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.text & ""
    End If
    
    Else
    crpt.ReplaceSelectionFormula "abs({tempLedgerRpt.Owner})<>0"
    End If


ElseIf Check2.value = 1 Then


    If cboStation1.text <> "" And txtAmount.text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.text <> "" And txtAmount.text <> "" Then
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.text & ""
    End If
    
    
    ElseIf cboStation1.text = "" And txtAmount.text <> "" Then
    
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.text & ""
    End If
    
    Else
    crpt.ReplaceSelectionFormula "abs({tempLedgerRpt.Owner})<>0"
    End If



End If

''======================================================================
''======================================================================

DSNNew

DoEvents
MsgBox ("View")
crpt.Connect = constr
crpt.Formulas(0) = "partyname='" & cboStation1.text & "'"
crpt.Formulas(1) = "ason='" & dateAson.value & "'"


crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.WindowShowSearchBtn = True
crpt.Action = 1

End Sub

Private Sub cmdprintalf_Click()
 If txtalfa.text = "" Then
    MsgBox "Please Enter Alphabet...", vbInformation
    Exit Sub
 End If
 
 Screen.MousePointer = vbHourglass
 'login.DSN
 CityWiseStatement
 Screen.MousePointer = vbDefault

End Sub

Private Sub cmdRefProm_Click()
For k_1 = 1 To vs.rows - 1
DoEvents
con.Execute "update DonnationMain set tobeupdate='' where dno=" & vs.TextMatrix(k_1, 1) & ""
vs.TextMatrix(k_1, 8) = "Click For Update..."
For k1 = 0 To 8
vs.Cell(flexcpBackColor, k_1, k1) = vbWhite
DoEvents
Next
DoEvents
DoEvents
Next
MsgBox "Data Refresh.....", vbInformation
End Sub

Private Sub cmdRepQty_Click()


Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application

Dim str_ As String



Screen.MousePointer = vbHourglass

If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double

Dim b1 As Boolean

b1 = False


c = 1
r = 1


row_ = 1
col_ = 1

xl.Columns("A:H").ColumnWidth = 12
J = 2
xlSheet.Cells(1, 1).value = "Bilty Return Status "

For I = 0 To vsop.rows - 1
    For J = 0 To vsop.Cols - 1
        If (col_ = 2 Or col_3) Then
           xlSheet.Cells(row_, col_).value = Format(vsop.TextMatrix(I, J), "dd/MM/yyyy")
        Else
           xlSheet.Cells(row_, col_).value = vsop.TextMatrix(I, J)
        End If
        col_ = col_ + 1
    Next
    row_ = row_ + 1
    col_ = 1
Next


Screen.MousePointer = vbDefault



End Sub

Private Sub cmdset_Click()
   
If RS.State = 1 Then RS.close
RS.Open "select * from pass where pass='" & cp & "'", con
If RS.EOF = True Then
   MsgBox "Enter Valid Password !!", vbInformation
   Exit Sub
End If
    
If Option2_donation.value = True Then
   
   If ch_din = "y" Then
      saveData
    Else
      
      frmPassword.Visible = True
      
      txtEnterPass.SetFocus
      MsgBox "Enter Password....", vbCritical
   End If
Else
   saveData
End If
    


   
End Sub
Sub saveData()
   
On Error GoTo ss:

Dim sysName_date As String

sysName_date = com_name & " - " & Date
   
   If MsgBox("Want to Set ?", vbQuestion + vbYesNo) = vbYes Then
        
   Screen.MousePointer = vbHourglass
   'cmdShow1.Visible = True
   Dim din As Integer
         
   If sales.value = True Then
        
        For J = 1 To vs.rows - 1
          
        If vs.TextMatrix(J, 5) <> "" Then
          
          If vs.TextMatrix(J, 5) = True Then
             If vs.TextMatrix(J, 5) = True Then din = 1
             con.Execute "update INVOICEA set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
             If vs.TextMatrix(J, 5) = False Then din = 0
             con.Execute "update INVOICEA set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        
        End If
        
          
          
        Next
        
 ElseIf Option2_app.value = True Then
        
        For J = 1 To vs.rows - 1
          
        If vs.TextMatrix(J, 5) <> "" Then
          
          If vs.TextMatrix(J, 5) = True Then
             If vs.TextMatrix(J, 5) = True Then din = 1
             con.Execute "update AppForm set bAuthorized=" & din & " where  appNO=" & vs.TextMatrix(J, 1) & ""
            Else
             If vs.TextMatrix(J, 5) = False Then din = 0
             con.Execute "update AppForm set bAuthorized=" & din & " where appNO=" & vs.TextMatrix(J, 1) & ""
          End If
        
        End If
          
          
        Next
        
 ElseIf Option2_agm.value = True Then
        
        For J = 1 To vs.rows - 1
            
            If vs.TextMatrix(J, 5) <> "" Then
              If vs.TextMatrix(J, 5) = True Then
                 If vs.TextMatrix(J, 5) = True Then din = 1
                 
                 con.Execute "update AgreementMain set bAuthorized=" & din & " where  AgmNo=" & vs.TextMatrix(J, 1) & ""
                Else
                 If vs.TextMatrix(J, 5) = False Then din = 0
                 con.Execute "update AgreementMain set bAuthorized=" & din & " where AgmNo=" & vs.TextMatrix(J, 1) & ""
              End If
            End If
            
        Next
        
 ElseIf Option2_donation.value = True Then
        
        For J = 1 To vs.rows - 1
            
            If vs.TextMatrix(J, 5) <> "" Then
              If vs.TextMatrix(J, 5) = True Then
                 If vs.TextMatrix(J, 5) = True Then din = 1
                 
                 con.Execute "update DonnationMain set date_sysname='" & sysName_date & "',UserName='" & UId & "',bAuthorized=" & din & ",status_='" & vs.TextMatrix(J, 7) & "' where  DNo=" & vs.TextMatrix(J, 1) & ""
                Else
                 If vs.TextMatrix(J, 5) = False Then din = 0
                 con.Execute "update DonnationMain set date_sysname='" & sysName_date & "',UserName='" & UId & "',bAuthorized=" & din & ",status_='" & vs.TextMatrix(J, 7) & "' where DNo=" & vs.TextMatrix(J, 1) & ""
              End If
            End If
            
        Next
    
 ElseIf Option_bookIssueSp.value = True Then
        
        For J = 1 To vs.rows - 1
        
        
        If vs.TextMatrix(J, 5) <> "" Then
           
          If vs.TextMatrix(J, 5) = "True" Then vs.TextMatrix(J, 5) = -1
          If vs.TextMatrix(J, 5) = True Then
             If vs.TextMatrix(J, 5) = -1 Then din = 1
             con.Execute "update INVOICEA_sp set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
             If vs.TextMatrix(J, 5) = False Then din = 0
             con.Execute "update INVOICEA_sp set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        End If
        
        
        Next
        
 ElseIf Option_bookRetSp.value = True Then
        
        For J = 1 To vs.rows - 1
        If vs.TextMatrix(J, 5) <> "" Then
          If vs.TextMatrix(J, 5) = True Then
             
             If vs.TextMatrix(J, 5) <> True Then
                If vs.TextMatrix(J, 5) = -1 Then din = 1
             Else
                din = 1
             End If
             
             con.Execute "update INVOICEA_spRet set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
             If vs.TextMatrix(J, 5) = False Then din = 0
             con.Execute "update INVOICEA_spRet set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        End If
        Next
   
  ElseIf credit.value = True Then
  
        For J = 1 To vs.rows - 1
          If vs.TextMatrix(J, 5) = "True" Then vs.TextMatrix(J, 5) = -1
          If vs.TextMatrix(J, 5) = True Then
            If vs.TextMatrix(J, 5) = -1 Then din = 1
            con.Execute "update CREDITA set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
            If vs.TextMatrix(J, 5) = False Then din = 0
            con.Execute "update CREDITA set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
  
  
  ElseIf cash.value = True Then
        
        For J = 1 To vs.rows - 1
        
        
         If vs.TextMatrix(J, 5) <> "" Then
          
          If vs.TextMatrix(J, 5) = True Then
             If vs.TextMatrix(J, 5) = True Then din = 1
             con.Execute "update casha set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
             If vs.TextMatrix(J, 5) = False Then din = 0
             con.Execute "update casha set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        
        End If
      
        Next
        
  
  ElseIf crdit.value = True Then
  
        For J = 1 To vs.rows - 1
          
          If vs.TextMatrix(J, 5) = True Then
            
            If vs.TextMatrix(J, 5) = False Then
               vs.TextMatrix(J, 5) = 0
            Else
               vs.TextMatrix(J, 5) = -1
            End If
            
            If vs.TextMatrix(J, 6) = False Then
               vs.TextMatrix(J, 6) = 0
            Else
               vs.TextMatrix(J, 6) = -1
            End If

            
            con.Execute "update cnf1a set bAuthorized=" & vs.TextMatrix(J, 5) & "  where " & stringyear & " and cnn=" & vs.TextMatrix(J, 1) & ""
            Else
            
            If vs.TextMatrix(J, 5) = False Then
               vs.TextMatrix(J, 5) = 0
            Else
               vs.TextMatrix(J, 5) = -1
            End If
            
            If vs.TextMatrix(J, 6) = False Then
               vs.TextMatrix(J, 6) = 0
            Else
               vs.TextMatrix(J, 6) = -1
            End If


            con.Execute "update cnf1a set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and cnn=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
  
  
  ElseIf dbit.value = True Then
  
        For J = 1 To vs.rows - 1
        
        
          If vs.TextMatrix(J, 5) = True Then
            
            If vs.TextMatrix(J, 5) = False Then
               vs.TextMatrix(J, 5) = 0
            Else
               vs.TextMatrix(J, 5) = -1
            End If
            

          
            con.Execute "update dnfa set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and dnn=" & vs.TextMatrix(J, 1) & ""
            Else
            
            If vs.TextMatrix(J, 5) = False Then
               vs.TextMatrix(J, 5) = 0
            Else
               vs.TextMatrix(J, 5) = -1
            End If
            
            con.Execute "update dnfa set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and dnn=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
   
   
  End If
   
   
End If
   

Screen.MousePointer = vbDefault

Exit Sub
ss:
MsgBox err.DESCRIPTION, vbInformation
Screen.MousePointer = vbDefault


End Sub
Private Sub cmdshow_Click()
      
Dim datecondition As String
Dim donationDate As String

Dim AgmDate As String
      
Screen.MousePointer = vbHourglass
cmdPrint_Pro.Visible = False
Command5_print.Visible = False
vs_promotion.Visible = False
'cmdTTrans.Visible = False

vs.Clear

Dim ClosingBal As Double

DoEvents
DoEvents

ClosingBal = 0


If Option2_donation.value = True Then
  cmdRefProm.Enabled = True
  cmdUpDatePromotion.Enabled = True
Else
 cmdRefProm.Enabled = False
 cmdUpDatePromotion.Enabled = False
End If



datecondition = "(convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & fromdate.value & "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & todate.value & "',103)"
donationDate = "(convert(smalldatetime,DDate,103)>=convert(smalldatetime,'" & fromdate.value & "',103) and convert(smalldatetime,DDate,103)<=convert(smalldatetime,'" & todate.value & "',103)"
      
AgmDate = "(convert(smalldatetime,Dates,103)>=convert(smalldatetime,'" & fromdate.value & "',103) and convert(smalldatetime,Dates,103)<=convert(smalldatetime,'" & todate.value & "',103)"
      
      
If sales.value = True Then
      
        
      
      If RS.State = 1 Then RS.close
      If txtParty.text = "" Then
        
        If All.value = True Then
           'RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from INVOICEA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") ORDER BY INVOICENO", con, adOpenStatic, adLockPessimistic
           Set RS = con.Execute("exec saleregisterAuth '2'")
        ElseIf autho.value = True Then
           'RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netAMOUNT,bauthorized from INVOICEA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=1 ORDER BY INVOICENO", con, adOpenStatic, adLockPessimistic
           Set RS = con.Execute("exec saleregisterAuth '" & 1 & "'")
        Else
           'RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from INVOICEA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=0 ORDER BY INVOICENO", con
           Set RS = con.Execute("exec saleregisterAuth '" & 0 & "'")
        End If
      
      Else
        If All.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from INVOICEA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' ORDER BY INVOICENO", con
        ElseIf autho.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from INVOICEA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=1 ORDER BY INVOICENO", con
        Else
         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from INVOICEA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=0 ORDER BY INVOICENO", con
        End If
      End If
      
      
        'Set rs_closing = con.Execute("exec saleregisterAuth '10'")
  
      
      
        vs.rows = 3
        ''vs.Cols = vs.Cols + 1
        If RS.EOF = False Then
        'vs.Rows = RS.RecordCount + 1
        For I = 1 To RS.RecordCount
        If RS.EOF = False Then
           'DoEvents
           
           rs_closing.MoveFirst
           rs_closing.Find "Subledger='" & RS.Fields(2).value & "'"
           If rs_closing.EOF = False Then
               vs.TextMatrix(I, 6) = Round(rs_closing(2), 0)
           End If

           
           vs.rows = vs.rows + 1
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           RS.MoveNext
          End If
           
        Next
      Else
           vs.Clear
           vs.rows = 2
      End If
      
End If

If credit.value = True Then

      If RS.State = 1 Then RS.close
      If txtParty.text = "" Then
        
        If All.value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized,Adj_YesNo from CREDITA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") ORDER BY INVOICENO", con
        ElseIf autho.value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized,Adj_YesNo from CREDITA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=1 ORDER BY INVOICENO", con
        Else
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized,Adj_YesNo from CREDITA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=0 ORDER BY INVOICENO", con
        End If
      
      Else
        If All.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized,Adj_YesNo from CREDITA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' ORDER BY INVOICENO", con
        ElseIf autho.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized,Adj_YesNo from CREDITA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=true ORDER BY INVOICENO", con
        Else
         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized,Adj_YesNo from CREDITA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=false ORDER BY INVOICENO", con
        End If
      End If
      If RS.EOF = False Then
        vs.rows = RS.RecordCount + 1
        For I = 1 To vs.rows - 1
        
           DoEvents
           DoEvents
           
           rs_closing.MoveFirst
           rs_closing.Find "Subledger='" & RS.Fields(2).value & "'"
           If rs_closing.EOF = False Then
               vs.TextMatrix(I, 6) = Round(rs_closing(2), 2)
           End If

           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           
           If RS!Adj_YesNo = "y" Then
            vs.Cell(flexcpBackColor, I, 0) = vbGreen
            vs.Cell(flexcpBackColor, I, 1) = vbGreen
            vs.Cell(flexcpBackColor, I, 2) = vbGreen
            vs.Cell(flexcpBackColor, I, 3) = vbGreen
            vs.Cell(flexcpBackColor, I, 4) = vbGreen
            vs.Cell(flexcpBackColor, I, 5) = vbGreen
            vs.Cell(flexcpBackColor, I, 6) = vbGreen
            DoEvents
          End If
          
           
           
           DoEvents
           DoEvents

           
           RS.MoveNext
        Next
      Else
           vs.Clear
           vs.rows = 2
      End If
End If

      
'==================

If cash.value = True Then
      
      If RS.State = 1 Then RS.close
      If txtParty.text = "" Then
        
        If All.value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,BAuthorized from CASHA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") ORDER BY INVOICENO", con
        ElseIf autho.value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=1   ORDER BY INVOICENO", con
        Else
           RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=0 ORDER BY INVOICENO", con
        End If
      
      Else
        If All.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' ORDER BY INVOICENO", con
        ElseIf autho.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=true ORDER BY INVOICENO", con
        Else
         RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,bauthorized from CASHA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=false ORDER BY INVOICENO", con
        End If
      End If
      
      
      
      
      If RS.EOF = False Then
        vs.rows = RS.RecordCount + 1
        For I = 1 To vs.rows - 1
           DoEvents
           
          
          
           rs_closing.MoveFirst
           rs_closing.Find "Subledger='" & RS.Fields(2).value & "'"
           If rs_closing.EOF = False Then
               vs.TextMatrix(I, 6) = Round(rs_closing(2), 2)
           End If

           
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           RS.MoveNext
           DoEvents

        Next
      Else
           vs.Clear
           vs.rows = 2
      End If
      
End If
      
      
'================================
If crdit.value = True Then
       


      datecondition = "(convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" & fromdate.value & "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & todate.value & "',103))"
            
      If RS.State = 1 Then RS.close
      If txtParty.text = "" Then
        
        If All.value = True Then
           RS.Open "select cnn,cnd,psld,na,BAuthorized,ReflectInAcc from cnf1a where " & stringyear & " and (" & datecondition & ") ORDER BY cnn", con
        ElseIf autho.value = True Then
           RS.Open "select cnn,cnd,psld,na,bauthorized,ReflectInAcc from cnf1a where " & stringyear & " and (" & datecondition & ") and bAuthorized=1 ORDER BY [cnn]", con
        Else
           RS.Open "select cnn,cnd,psld,na,bauthorized,ReflectInAcc from cnf1a where " & stringyear & " and (" & datecondition & ") and bAuthorized=0 ORDER BY cnn", con
        End If
      
      Else
        If All.value = True Then
         RS.Open "select cnn,cnd,psld,na,bauthorized,ReflectInAcc from cnf1a where " & stringyear & " and (" & datecondition & ") and psld='" & txtParty.text & "' ORDER BY cnn", con
        ElseIf autho.value = True Then
         RS.Open "select cnn,cnd,psld,na,bauthorized,ReflectInAcc from cnf1a where " & stringyear & " and (" & datecondition & ") and psld='" & txtParty.text & "' and bAuthorized=true ORDER BY cnn", con
        Else
         RS.Open "select cnn,cnd,psld,na,bauthorized,ReflectInAcc from cnf1a where " & stringyear & " and (" & datecondition & ") and psld='" & txtParty.text & "' and bAuthorized=false ORDER BY cnn", con
        End If
      End If
      
      
      
      
      If RS.EOF = False Then
        vs.rows = RS.RecordCount + 1
        For I = 1 To vs.rows - 1
           DoEvents
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           vs.TextMatrix(I, 6) = RS.Fields(5).value & ""
           
           
           rs_closing.MoveFirst
           rs_closing.Find "Subledger='" & RS.Fields(2).value & "'"
           If rs_closing.EOF = False Then
               vs.TextMatrix(I, 6) = Round(rs_closing(2), 2)
           End If

           
           RS.MoveNext
           DoEvents
        Next
      Else
           vs.Clear
           vs.rows = 2
      End If

End If

If Option2_app.value = True Then

      If RS.State = 1 Then RS.close
      If txtParty.text = "" Then
        
        If All.value = True Then
           RS.Open "select distinct AppNo,AppDate,School_PartyName,NetAmt,BAuthorized from AppForm  ORDER BY appno", con
        ElseIf autho.value = True Then
           RS.Open "select distinct AppNo,AppDate,School_PartyName,NetAmt,BAuthorized from AppForm where BAuthorized=1 ORDER BY AppNo", con
        Else
           RS.Open "select distinct AppNo,AppDate,School_PartyName,NetAmt,BAuthorized from AppForm where BAuthorized=0 ORDER BY AppNo", con
        End If
      
      Else
         'If All.value = True Then
         'RS.Open "select dnn,dnd,psld,na,bauthorized from dnfa where " & stringyear & " and (" & datecondition & ") and psld='" & txtParty.Text & "' ORDER BY Dnn", con
         'ElseIf autho.value = True Then
         'RS.Open "select dnn,dnd,psld,na,bauthorized from dnfa where " & stringyear & " and (" & datecondition & ") and psld='" & txtParty.Text & "' and bAuthorized=true ORDER BY Dnn", con
         'Else
         'RS.Open "select dnn,dnd,psld,na,bauthorized from dnfa where " & stringyear & " and (" & datecondition & ") and psld='" & txtParty.Text & "' and bAuthorized=false ORDER BY Dnn", con
         'End If
      End If
      
      
      
      
      If RS.EOF = False Then
        vs.rows = RS.RecordCount + 1
        For I = 1 To vs.rows - 1
           DoEvents
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value & ""
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           
           
'           rs_closing.MoveFirst
'           rs_closing.Find "Subledger='" & RS.Fields(2).value & "'"
'           If rs_closing.EOF = False Then
'               vs.TextMatrix(I, 6) = Round(rs_closing(2), 2)
'           End If

           
           RS.MoveNext
           DoEvents
        Next
      Else
           vs.Clear
           vs.rows = 2
      End If


'Exit Sub

End If
      
If dbit.value = True Then
       
      datecondition = "(convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" & fromdate.value & "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & todate.value & "',103))"

      If RS.State = 1 Then RS.close
      If txtParty.text = "" Then
        
        If All.value = True Then
           RS.Open "select dnn,dnd,psld,na,BAuthorized from dnfa where " & stringyear & " and (" & datecondition & ") ORDER BY Dnn", con
        ElseIf autho.value = True Then
           RS.Open "select dnn,dnd,psld,na,bauthorized from dnfa where " & stringyear & " and (" & datecondition & ") and bAuthorized=1 ORDER BY Dnn", con
        Else
           RS.Open "select dnn,dnd,psld,na,bauthorized from dnfa where " & stringyear & " and (" & datecondition & ") and bAuthorized=0 ORDER BY Dnn", con
        End If
      
      Else
        If All.value = True Then
         RS.Open "select dnn,dnd,psld,na,bauthorized from dnfa where " & stringyear & " and (" & datecondition & ") and psld='" & txtParty.text & "' ORDER BY Dnn", con
        ElseIf autho.value = True Then
         RS.Open "select dnn,dnd,psld,na,bauthorized from dnfa where " & stringyear & " and (" & datecondition & ") and psld='" & txtParty.text & "' and bAuthorized=true ORDER BY Dnn", con
        Else
         RS.Open "select dnn,dnd,psld,na,bauthorized from dnfa where " & stringyear & " and (" & datecondition & ") and psld='" & txtParty.text & "' and bAuthorized=false ORDER BY Dnn", con
        End If
      End If
      
      
      
      
      If RS.EOF = False Then
        vs.rows = RS.RecordCount + 1
        For I = 1 To vs.rows - 1
           DoEvents
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           
           rs_closing.MoveFirst
           rs_closing.Find "Subledger='" & RS.Fields(2).value & "'"
           If rs_closing.EOF = False Then
               vs.TextMatrix(I, 6) = Round(rs_closing(2), 2)
           End If

           
           RS.MoveNext
           DoEvents
        Next
      Else
           vs.Clear
           vs.rows = 2
      End If

End If
      

'=====Invoice_sp============================================================
If Option_bookIssueSp.value = True Then
      
      If RS.State = 1 Then RS.close
      If txtParty.text = "" Then
        
        If All.value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,AgentName,netAMOUNT,bauthorized from INVOICEA_sp where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") ORDER BY INVOICENO", con
        ElseIf autho.value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,AgentName,netAMOUNT,bauthorized from INVOICEA_sp where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=1 ORDER BY INVOICENO", con
        Else
           RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,bauthorized from INVOICEA_sp where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=0 ORDER BY INVOICENO", con
        End If
      
      Else
        If All.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,bauthorized from INVOICEA_sp where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' ORDER BY INVOICENO", con
        ElseIf autho.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,bauthorized from INVOICEA_sp where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=1 ORDER BY INVOICENO", con
        Else
         RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,bauthorized from INVOICEA_sp where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=0 ORDER BY INVOICENO", con
        End If
      End If
      
      
      
      If RS.EOF = False Then
        vs.rows = RS.RecordCount + 1
        For I = 1 To vs.rows - 1
           DoEvents
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           RS.MoveNext
           DoEvents
        Next
      Else
           vs.Clear
           vs.rows = 2
      End If
      
End If
      

'=================================================================

'=====Invoice_sp============================================================
If Option_bookRetSp.value = True Then
      
      If RS.State = 1 Then RS.close
      If txtParty.text = "" Then
        
        If All.value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,agentname,netAMOUNT,bauthorized from INVOICEA_spRet where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") ORDER BY INVOICENO", con
        ElseIf autho.value = True Then
           RS.Open "select INVOICENO,INVOICEDATE,agentname,netAMOUNT,bauthorized from INVOICEA_spRet where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=1 ORDER BY INVOICENO", con
        Else
           RS.Open "select INVOICENO,INVOICEDATE,agentname,netamount,bauthorized from INVOICEA_spRet where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and bAuthorized=0 ORDER BY INVOICENO", con
        End If
      
      Else
        If All.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,agentname,netamount,bauthorized from INVOICEA_spRet where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' ORDER BY INVOICENO", con
        ElseIf autho.value = True Then
         RS.Open "select INVOICENO,INVOICEDATE,agentname,netamount,bauthorized from INVOICEA_spRet where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=1 ORDER BY INVOICENO", con
        Else
         RS.Open "select INVOICENO,INVOICEDATE,agentname,netamount,bauthorized from INVOICEA_spRet where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0) and " & datecondition & ") and SUBLEDGER='" & txtParty.text & "' and bAuthorized=0 ORDER BY INVOICENO", con
        End If
      End If
      
      
      
      If RS.EOF = False Then
        vs.rows = RS.RecordCount + 1
        For I = 1 To vs.rows - 1
           DoEvents
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           RS.MoveNext
           DoEvents
        Next
      Else
           vs.Clear
           vs.rows = 2
      End If
      
End If
      


'=================================================================
'=====Invoice_sp============================================================
If Option2_donation.value = True Then
      
      'If ch_din = "n" Then
      '   Screen.MousePointer = vbDefault
      '   Exit Sub
      'End If
      
      cmdPrint_Pro.Visible = True
      Command5_print.Visible = True
      'cmdTTrans.Visible = True
      
      If RS.State = 1 Then RS.close
        
        If All.value = True Then
           RS.Open "select DNo,DDate,ScName + ':' + scid,(NetBalance+AdvAmt) as NETAMOUNT,bauthorized,status_,RoundOfAAmt,RoundOfAAmt_New,RoundOfAAmt from DonnationMain where " & donationDate & ") ORDER BY DNo", con
        ElseIf autho.value = True Then
           RS.Open "select DNo,DDate,ScName+ ':' + scid,(NetBalance+AdvAmt) as NETAMOUNT,bauthorized,status_,RoundOfAAmt,RoundOfAAmt_New,RoundOfAAmt from DonnationMain where " & donationDate & ") and bAuthorized=1 ORDER BY DNo", con
        Else
           RS.Open "select DNo,DDate,ScName+ ':' + scid,(NetBalance+AdvAmt) as NETAMOUNT,bauthorized,status_,RoundOfAAmt,RoundOfAAmt_New,RoundOfAAmt from DonnationMain where " & donationDate & ") and bAuthorized=0 ORDER BY DNo", con
        End If
      
      
      vs.Cols = 9
      
      
      If RS.EOF = False Then
        vs.rows = RS.RecordCount + 1
        For I = 1 To vs.rows - 1
           DoEvents
           
           If RS.Fields(0).value = 149 Then
              MsgBox "a"
           End If
           
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           vs.TextMatrix(I, 7) = RS.Fields(5).value & ""
           
           If (IsNull(RS.Fields(8).value) Or RS.Fields(8).value = "") Then
              vs.TextMatrix(I, 8) = "Click For Update..."
           Else
              vs.TextMatrix(I, 8) = "Round Of Amount : " & RS.Fields(8).value
              If RS.Fields(8).value < 0 Then
              For k1 = 0 To 8
                  vs.Cell(flexcpBackColor, I, k1) = vbGreen
              DoEvents
              Next
              End If

           End If
           'vs.TextMatrix(I, 9) = RS.Fields(7).value & ""
           
           RS.MoveNext
           DoEvents
        Next
      Else
           vs.Clear
           vs.rows = 2
      End If
      
      
      
End If
      

'=================================================================
If Option2_agm.value = True Then

    Set RS = New ADODB.Recordset

    If All.value = True Then
       RS.Open "select distinct AgmNo,Dates as AgmDate,PName,ExpSale,BAuthorized from AgreementMain  ORDER BY AgmNo", con
    ElseIf autho.value = True Then
       RS.Open "select distinct AgmNo,Dates as AgmDate,PName,ExpSale,BAuthorized from AgreementMain where BAuthorized=1 ORDER BY AgmNo", con
    Else
       RS.Open "select distinct AgmNo,Dates as AgmDate,PName,ExpSale,BAuthorized from AgreementMain where BAuthorized=0 ORDER BY AgmNo", con
    End If
    
    
     If RS.EOF = False Then
        vs.rows = RS.RecordCount + 1
        For I = 1 To vs.rows - 1
           DoEvents
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           RS.MoveNext
           DoEvents
        Next
      Else
           vs.Clear
           vs.rows = 2
      End If

End If



 vsIni
 Screen.MousePointer = vbDefault
 
End Sub
Sub SearchFa_blueprint()
      
      ''BluePrint
      
      If RS.State = 1 Then RS.close
      RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,BAA,t2 from INVOICEA_blue where " & stringyear & " and SUBLEDGER='" & cboParty.text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
      If RS.EOF = False Then
        vs1.rows = (vs1.rows + RS.RecordCount)
        For I = I To vs1.rows - 1
        If RS.EOF = False Then
           vs1.TextMatrix(I, 0) = "I"
           vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
           vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
           If IsNull(RS.Fields("t2").value) Then
              vs1.TextMatrix(I, 3) = "Invoice Sales"
           Else
              vs1.TextMatrix(I, 3) = "Invoice Sales" & RS.Fields("t2").value & " " & "DS"
           End If
           vs1.TextMatrix(I, 4) = Format(RS.Fields("netamount").value, "0.00")
           vs1.TextMatrix(I, 5) = Format(RS.Fields("BAA").value, "0.00")
           
           
            RS.MoveNext
         End If
        Next
      End If

  
    
    vs1.FormatString = "^Bill Type|^Bill|^Date|<Description|>Dr|>Cr"
    setWidth

End Sub
Sub SearchFa()
      Dim s10 As String
      
      con.Execute ("exec UpdateSDiscount 'invoiceno'")
     
      If RS.State = 1 Then RS.close
      RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,BAA,t2,todid,todDate,BILTYNO,SCName,BUNDLES,sdiscount,app_add from INVOICEA where " & stringyear & " and SUBLEDGER='" & cboParty.text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
      If RS.EOF = False Then
        vs1.rows = (vs1.rows + RS.RecordCount)
        For I = I To vs1.rows - 1
        
        s10 = ""
        
        If RS.EOF = False Then
           vs1.TextMatrix(I, 0) = "I"
           vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
           vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
           
           If Not IsNull(RS!sdiscount) Then
           If Val(RS!sdiscount) > 0 Then
              s10 = "(DS -" & RS!sdiscount & ")"
           End If
           End If
           
           If Not IsNull(RS!biltyno) Then
           If Val(RS!biltyno) > 0 Then
           If s10 = "" Then
                s10 = "Bilty No-" & RS!biltyno
              Else
                s10 = s10 & ", Bilty No-" & RS!biltyno
           End If
           End If
           End If
           
           If Not IsNull(RS!bundles) Then
           If Len(RS!bundles) > 0 Then
            If s10 = "" Then
               s10 = "BUNDLES-" & RS!bundles
            Else
               s10 = s10 & ", BUNDLES - " & RS!bundles
            End If
           End If
           End If

           
           If Not IsNull(RS!scname) Then
           If Len(RS!scname) > 0 Then
              If s10 = "" Then
                 s10 = "SCHOOL-" & RS!scname
              Else
                 s10 = s10 & ",SCHOOL-" & RS!scname
              End If
           End If
           End If
           
           
           If IsNull(RS.Fields("t2").value) Then
              
              vs1.TextMatrix(I, 3) = "Invoice Sales  " & s10
           Else
              vs1.TextMatrix(I, 3) = "Invoice Sales" & RS.Fields("t2").value & " " & "DS"
           End If
           vs1.TextMatrix(I, 4) = Format(RS.Fields("netamount").value, "0.00")
           vs1.TextMatrix(I, 5) = Format(RS.Fields("BAA").value, "0.00")
           
           If Not IsNull(RS!todid) Then
              vs1.TextMatrix(I, 10) = RS.Fields("todid").value & "-" & RS.Fields("toddate").value
           End If
           

           
           
           RS.MoveNext
         End If
        Next
      End If

    ''================
   
     If RS.State = 1 Then RS.close
     RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,baa,todid,toddate,SCName,BILTYNO,bundles from CREDITA where " & stringyear & " and SUBLEDGER='" & cboParty.text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
     If RS.EOF = False Then
        vs1.rows = vs1.rows + RS.RecordCount
        For I = I To vs1.rows - 1
         
        s10 = ""
         
        If RS.EOF = False Then
        
           
           If Not IsNull(RS!biltyno) Then
           If Val(RS!biltyno) > 0 Then
              s10 = "Bilty No-" & RS!biltyno
           End If
           End If
           
           
           
           If Not IsNull(RS!bundles) Then
           If Len(RS!bundles) > 0 Then
            If s10 = "" Then
               s10 = "BUNDLES - " & RS!bundles
            Else
               s10 = s10 & ", BUNDLES-" & RS!bundles
            End If
           End If
           End If

           
           
           If Not IsNull(RS!scname) Then
           If Len(RS!scname) > 0 Then
              If s10 = "" Then
                 s10 = "SCHOOL-" & RS!scname
              Else
                 s10 = s10 & ",SCHOOL-" & RS!scname
              End If
           End If
           End If
           
        
         vs1.TextMatrix(I, 0) = "CI"
         vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
         vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
         vs1.TextMatrix(I, 3) = "Cr. Note Item " & s10
         vs1.TextMatrix(I, 4) = Format(RS.Fields("baa").value, "0.00")
         vs1.TextMatrix(I, 5) = Format(RS.Fields("netamount").value, "0.00")
         If Not IsNull(RS!todid) Then
            vs1.TextMatrix(I, 10) = RS.Fields("todid").value & "-" & RS.Fields("toddate").value
         End If

         RS.MoveNext
       End If
    Next
    End If
    If RS.State = 1 Then RS.close
    RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,baa,t2,sdiscount from CASHA where  " & stringyear & " and SUBLEDGER='" & cboParty.text & "'  and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
    If RS.EOF = False Then
     vs1.rows = vs1.rows + RS.RecordCount
     For I = I To vs1.rows - 1
    If RS.EOF = False Then
      vs1.TextMatrix(I, 0) = "C/M"
      vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
      vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
      If IsNull(RS.Fields("t2").value) Then
         vs1.TextMatrix(I, 3) = "Cash Memo"
      Else
         vs1.TextMatrix(I, 3) = "Cash Memo" & " " & RS.Fields("t2").value & " DS"
      End If
      
      If vs1.TextMatrix(I, 3) = "" Then
         If Not IsNull(RS!sdiscount) Then
            vs1.TextMatrix(I, 3) = "(DS-" & RS!sdiscount & ")"
         End If
      Else
         If Not IsNull(RS!sdiscount) Then
            vs1.TextMatrix(I, 3) = vs1.TextMatrix(I, 3) & " (DS-" & RS!sdiscount & ")"
         End If

      End If
      
      
      vs1.TextMatrix(I, 4) = Format(RS.Fields("netamount").value, "0.00")
      vs1.TextMatrix(I, 5) = Format(RS.Fields("baa").value, "0.00")
      RS.MoveNext
    End If
    Next
    End If

 '===================
    If RS.State = 1 Then RS.close
    RS.Open "select cnn,cnd,na,todid,toddate,CNCategory from Cnf1a where  " & stringyear & " and psld='" & cboParty.text & "'", con
    If RS.EOF = False Then
     vs1.rows = vs1.rows + RS.RecordCount
     For I = I To vs1.rows - 1
      
     s10 = ""
      
     If rs1.State = 1 Then rs1.close
     rs1.Open "select narr from CreditNotDet where cnn='" & RS.Fields("cnn").value & "'", con
     While rs1.EOF = False
        If s10 = "" Then
           s10 = rs1!NARR
        Else
           s10 = s10 & ", " & rs1!NARR
        End If
     rs1.MoveNext
     Wend
      
     If RS.EOF = False Then
    
      vs1.TextMatrix(I, 0) = "CN"
      vs1.TextMatrix(I, 1) = RS.Fields("cnn").value
      vs1.TextMatrix(I, 2) = RS.Fields("cnd").value
      vs1.TextMatrix(I, 3) = "Credit Note -" & s10
      vs1.TextMatrix(I, 5) = Format(RS.Fields("na").value, "0.00")
      vs1.TextMatrix(I, 4) = 0
      
      If Not IsNull(RS!todid) Then
         If RS!CNCategory = "Adjustment" Then
            vs1.TextMatrix(I, 10) = RS.Fields("todid").value & "-" & RS.Fields("toddate").value
         End If
      End If

      
      RS.MoveNext
    
    End If
    
    Next
    End If
     
     
     
    '===================
    If RS.State = 1 Then RS.close
    RS.Open "select dnn,dnd,psld,na,n from dnfa where  " & stringyear & " and psld='" & cboParty.text & "'", con
    If RS.EOF = False Then
     vs1.rows = vs1.rows + RS.RecordCount
     For I = I To vs1.rows - 1
     
     s10 = ""
    
     If RS.EOF = False Then
    
      If (Not IsNull(RS!n) Or RS!n <> "") Then s10 = "" & RS!n
     
     '----------------------------------
       If rs1.State = 1 Then rs1.close
       rs1.Open "select narr from debitNotDet where dnn='" & RS.Fields("dnn").value & "'", con
       If rs1.EOF = False Then s10 = ""
       While rs1.EOF = False
         If s10 = "" Then
           s10 = rs1!NARR
         Else
           s10 = s10 & ", " & rs1!NARR
         End If
        rs1.MoveNext
       Wend
     '------------------------------------
      
        
      vs1.TextMatrix(I, 0) = "DN"
      vs1.TextMatrix(I, 1) = RS.Fields("dnn").value
      vs1.TextMatrix(I, 2) = RS.Fields("dnd").value
      vs1.TextMatrix(I, 3) = "Debit Note " & s10
      vs1.TextMatrix(I, 4) = Format(RS.Fields("na").value, "0.00")
      vs1.TextMatrix(I, 5) = 0
      RS.MoveNext
    End If
    Next
    End If
    
  
    
    vs1.FormatString = "^Bill Type|^Bill|^Date|<Description|>Dr|>Cr"
    setWidth
End Sub
Sub CityWiseStatement()



Dim op, drcr
Dim s As String
Dim sts As String

s = ""
Dim dist, Party1, RepName As String
Dim rs1 As New ADODB.Recordset
Dim rs1_dist As New ADODB.Recordset
Dim rs1_rpt As New ADODB.Recordset
Dim rs1_rep As New ADODB.Recordset

DSNNew
       
If rs1_dist.State = 1 Then rs1_dist.close
rs1_dist.Open "select DISTCODE,subledger,DESCFORINVOICE,states,ADDRESS3 from SLEDGER where (gledger='SUNDRY DEBTORS') group by DISTCODE,subledger,DESCFORINVOICE,states,ADDRESS3", con, adOpenDynamic, adLockOptimistic
      
con.Execute "delete from templedger1 where  " & stringyear & " and userid='" & UserName & "' and rptype='" & cboStation & "'"
con.Execute "delete from templedger1 where userid is null"


If RS.State = 1 Then RS.close

If rs1_rpt.State = 1 Then rs1_rpt.close
rs1_rpt.Open "select max(rptid) from tempLedger1", con, adOpenDynamic, adLockOptimistic
If IsNull(rs1_rpt(0)) Then
rptid = 9999
Else
rptid = rs1_rpt(0) + 1
End If

       
If cboStation.text <> "" And txtalfa.text = "" Then
'=====================================================================================
For I = 0 To cboPartyList.ListCount - 1
If cboPartyList.Selected(I) = True Then
If s = "" Then
  s = "SUBLEDGER " & " = " & "'" & cboPartyList.List(I) & "'"
Else
  s = s & " or " & "SUBLEDGER " & " = " & "'" & cboPartyList.List(I) & "'"
End If
End If
Next
If Option3_dist.value = True Then
       
       If s = "" Then
       If RS.State = 1 Then RS.close
        RS.Open "select subledger from SLEDGER where " & stringyear & " and DISTCODE = '" & cboStation.text & "'", con
       Else
        If RS.State = 1 Then RS.close
        RS.Open "select subledger from SLEDGER where " & stringyear & " and " & s, con
       End If
       


    ElseIf Option2_state.value = True Then

      If s = "" Then
        If RS.State = 1 Then RS.close
        If cboStation.text = "ALL" Then
        RS.Open "select subledger from SLEDGER where (gledger='SUNDRY DEBTORS')", con
        Else
        RS.Open "select subledger from SLEDGER where ((gledger='SUNDRY DEBTORS') AND states = '" & cboStation.text & "')", con
        End If
        
       Else
        If RS.State = 1 Then RS.close
        RS.Open "select subledger from SLEDGER where (gledger='SUNDRY DEBTORS') and " & s, con
       End If

      

    ElseIf Option4_rep.value = True Then

       If s = "" Then
        If cboStation.text = "ALL" Then
            If RS.State = 1 Then RS.close
            RS.Open "select distinct(subledger) from SLEDGER", con
        Else
            If RS.State = 1 Then RS.close
            RS.Open "select distinct(subledger) from INVOICEA where AgentName = '" & cboStation.text & "'", con
        End If
       
       Else
        If RS.State = 1 Then RS.close
        RS.Open "select subledger from SLEDGER where (gledger='SUNDRY DEBTORS') and " & s, con
       End If
       

    End If
    
    
'=====================================================================================
ElseIf txtalfa.text <> "" And cboStation.text = "" Then
       RS.Open "select subledger from SLEDGER where ((gledger='SUNDRY DEBTORS') and Subledger like '" + Trim(txtalfa.text) + "%')", con
End If

While RS.EOF = False
rs1_dist.MoveFirst
rs1_dist.Find "subledger='" & RS.Fields(0).value & "'"
If rs1_dist.EOF = False Then
dist = Trim(rs1_dist.Fields("ADDRESS3").value)
sts = rs1_dist.Fields("states").value
Party1 = rs1_dist.Fields("DESCFORINVOICE").value
End If
If rs1.State = 1 Then rs1.close
rs1.Open "SELECT op,drcr from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'", con
If Not IsNull(rs1.Fields(0).value) Then
op = Val(rs1.Fields(0).value)
drcr = rs1.Fields(1).value
Else
op = 0
End If

con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype,rptid,rptype,setupid,fyear,district,userid,states,Party1)  SELECT op,'" & drcr & "',subledger,'Opening'," & rptid & ",'" & cboStation & "','" & setupid & "','" & session & "','" & dist & "','" & UserName & "','" & sts & "','" & Party1 & "' from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'   group by op,subledger,drcr HAVING  op <> 0;"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales Bilty No-' + BILTYNO + ',Bundle-' + bundles ,netamount,BAA,SUBLEDGER,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "',scname  from INVOICEA where convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,SUBLEDGER,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "',scname from CREDITA where convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "' from CASHA where  convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"

con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "' from dnfa where  convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and psld='" & RS.Fields(0).value & "'"
'''con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) select dnd,'DN',dnfa.dnn,'Debit Note',na,'0',PSLD,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "' from dnfa LEFT OUTER JOIN dbo.debitNotDet ON dbo.dnfa.DNN = dbo.debitnotDet.DNN where  (debitNotDet.RepName is null and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and psld='" & RS.Fields(0).value & "')"
'''con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) select dnd,'DN',dnfa.dnn,'Debit Note',debitNotDet.amount,'0',PSLD,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "' from dnfa LEFT OUTER JOIN dbo.debitNotDet ON dbo.dnfa.DNN = dbo.debitnotDet.DNN where  (debitNotDet.RepName is not null and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and psld='" & RS.Fields(0).value & "')"

con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) Select cnd,'CN',cnn,'Credit Note ' + desc_ ,na,'0',psld,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "' from Cnf1a where  convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and psld='" & RS.Fields(0).value & "'"
'''con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) Select cnd,'CN',CNF1A.cnn,'Credit Note ' + desc_ ,na,'0',psld,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "' from Cnf1a LEFT OUTER JOIN dbo.CreditNotDet ON dbo.CNF1A.CNN = dbo.CreditNotDet.CNN where (CreditNotDet.RepName is null and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and psld='" & RS.Fields(0).value & "')"
'''con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) Select cnd,'CN',CNF1A.cnn,'Credit Note ' + desc_ ,CreditNotDet.amount,'0',psld,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "' from Cnf1a LEFT OUTER JOIN dbo.CreditNotDet ON dbo.CNF1A.CNN = dbo.CreditNotDet.CNN where  (CreditNotDet.RepName is not null and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and psld='" & RS.Fields(0).value & "')"

con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) Select dates,'J',Recno,Particullar,Dr,CR,PartyName,fyear,setupid,'" & UserName & "','" & dist & "'," & rptid & ",'" & cboStation & "','" & sts & "','" & Party1 & "','" & RepName & "' from ReceiveIssueParty where convert(smalldatetime,DATEs,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and " & stringyear & " and PartyName='" & RS.Fields(0).value & "' and firm='" & firm & "' order by dates,recno"


If op <> 0 Then
con.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "',userid='" & UserName & "' where " & stringyear & " and party = '" + RS.Fields(0).value + "' and Billtype<>'Opening'"
End If
rptid = rptid + 1
RS.MoveNext
Wend


       
If Option4_rep = False Then

If Check_rep_billwise.value = 1 Then
   con.Execute "Update templedger1 set rptName='PartyLedger.rpt'"
   popupvalue4 = "PartyLedger.rpt"
Else
   con.Execute "Update templedger1 set rptName='PartyLedgerRep.rpt'"
   popupvalue4 = "PartyLedgerRep.rpt"
End If
  
Else
  con.Execute "Update templedger1 set rptName='PartyLedgerRep.rpt'"
  popupvalue4 = "PartyLedgerRep.rpt"
End If

If Check3_AmountTobecollect.value = 1 Then
  con.Execute "Update templedger1 set AspectedAmt='1'"
End If
DoEvents


con.Execute ("exec PartyStateMent")
 
 
 
 If MsgBox("Want to Send Mail", vbYesNo) = vbNo Then
 
 
 If Option4_rep = False Then
    
 If cboStation.text = "ALL" Then
 
    crpt.Reset
    crpt.ReportFileName = rptPath & "\PartyLedgerBillWiseAll.rpt"
    crpt.Connect = constr
    crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UserName & "' and {templedgerrpt.rptype}='" & cboStation & "'"
    crpt.WindowShowPrintSetupBtn = True
    crpt.Formulas(0) = "partyname='" & cboStation.text & "'"
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    crpt.Action = 1
    
    
 Else
 
 If txtalfa.text <> "" Then
    crpt.Reset
    crpt.ReportFileName = rptPath & "\PartyLedgerALF.rpt"
    crpt.Connect = constr
    crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UserName & "' and {templedgerrpt.rptype}='" & cboStation & "'"
    crpt.WindowShowPrintSetupBtn = True
    crpt.Formulas(0) = "partyname='" & cboStation.text & "'"
    'crpt.Formulas(1) = "dateason='" & txt_ason.value & "'"
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    crpt.Action = 1
 Else
    crpt.Reset
    If Check_rep_billwise.value = 1 Then
     crpt.ReportFileName = rptPath & "\PartyLedger.rpt"
    Else
     crpt.ReportFileName = rptPath & "\PartyLedgerRep.rpt"
    End If
    
    crpt.Connect = constr
    crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UserName & "' and {templedgerrpt.rptype}='" & cboStation & "'"
    crpt.WindowShowPrintSetupBtn = True
    crpt.Formulas(0) = "partyname='" & cboStation.text & "'"
    'crpt.Formulas(1) = "dateason='" & txt_ason.value & "'"
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    crpt.Action = 1
 
 
 
 End If
 
 End If
 
 Else
 
 
 If cboStation.text = "ALL" Then
    
  If Check_rep_billwise.value = 0 Then
  
    crpt.Reset
    crpt.ReportFileName = rptPath & "\PartyLedgerRepAll.rpt"
    crpt.Connect = constr
    crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UserName & "' and {templedgerrpt.rptype}='" & cboStation & "'"
    crpt.WindowShowPrintSetupBtn = True
    crpt.Formulas(0) = "partyname='" & cboStation.text & "'"
    crpt.Formulas(1) = "dateason='" & txt_ason.value & "'"
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    crpt.Action = 1
 
 Else
    crpt.Reset
    crpt.ReportFileName = rptPath & "\PartyLedgerRepBillWise.rpt"
    crpt.Connect = constr
    crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UserName & "' and {templedgerrpt.rptype}='" & cboStation & "'"
    crpt.WindowShowPrintSetupBtn = True
    crpt.Formulas(0) = "partyname='" & cboStation.text & "'"
    'crpt.Formulas(1) = "dateason='" & txt_ason.value & "'"
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    crpt.Action = 1
 
 End If
 
 Else
    
 If Check_rep_billwise.value = 0 Then
 
    crpt.Reset
    crpt.ReportFileName = rptPath & "\PartyLedgerRep.rpt"
    crpt.Connect = constr
    crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UserName & "' and {templedgerrpt.rptype}='" & cboStation & "'"
    crpt.WindowShowPrintSetupBtn = True
    crpt.Formulas(0) = "partyname='" & cboStation.text & "'"
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    crpt.WindowShowRefreshBtn = True
    crpt.Action = 1
 
 Else
 
    crpt.Reset
    crpt.ReportFileName = rptPath & "\PartyLedgerRepBillWise.rpt"
    crpt.Connect = constr
    crpt.ReplaceSelectionFormula "{templedgerrpt.Userid}='" & UserName & "' and {templedgerrpt.rptype}='" & cboStation & "'"
    crpt.WindowShowPrintSetupBtn = True
    crpt.Formulas(0) = "partyname='" & cboStation.text & "'"
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    crpt.Action = 1
 
 
 End If
 
 End If
 
 
 End If
 
 Else
   
   Screen.MousePointer = vbDefault
   popupvalue5 = rptid
   frmSendMail.Show 1
   
 End If
 
End Sub
Sub showData()

Dim contr As New ADODB.Connection
Dim fillVs As New ADODB.Recordset

Dim dr, CR As Double
Dim notmatch As Boolean

Dim bb As Boolean
Screen.MousePointer = vbHourglass
'
'-------------------------------------------------------------------------
Dim db, db1 As String

Dim dt1, dt2 As String
Dim rs_Acc As New ADODB.Recordset

dt1 = Format(fromdate.value, "MM/dd/yyyy")
dt2 = Format(dateAson.value, "MM/dd/yyyy")

Set rs_Acc = New ADODB.Recordset
Set rs_Acc = New ADODB.Recordset
Set rs_Acc = con.Execute("exec spBalanceSheet_Mearging_ch '" & dt1 & "','" & dt2 & "','SUNDRY DEBTORS'")

db = ""
db1 = ""

If Check3_ledgerClosingTrans.value = 1 Then

    I = Right(session, 2) + 1
    J = I - 1
    db1 = J & "" & I
    
    
    db = "chitradata_" & db1
    Set contr = New ADODB.Connection
    
    If LCase(server_) = "server" Then
       contr.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db & "; UID=" & sql_user & "; PWD=" & sql_pass
       contr.Open
    Else
       contr.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & db & "; UID=; PWD=;"
       contr.Open
    End If
    
    bb = True
    contr.Execute "update sledger set OP=" & 0 & ",drcr=''"

End If

'-------------------------------------------------------------------------


If MsgBox("Want To Show Balance", vbInformation + vbYesNo) <> vbYes Then
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Dim op, drcr
Dim rs1 As New ADODB.Recordset
con.Execute "delete from templedger2"

'CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER from INVOICEA WHERE " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER from INVOICEA WHERE ((netamount-BAA)>0 or (netamount-BAA)<0)"

'CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,SUBLEDGER from CREDITA WHERE " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,SUBLEDGER from CREDITA WHERE ((netamount-BAA)>0 or (netamount-BAA)<0)"

'CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER from CASHA where " & stringyear & " and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER from CASHA where ((netamount-BAA)>0 or (netamount-BAA)<0)"

'CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa where " & stringyear & ""

con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a"
'CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a where " & stringyear & " "

'CON.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty where " & stringyear & " and firm='" & firm & "'"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty"
DoEvents
DoEvents
'CON.Execute "update SLEDGER set Owner=0 where " & stringyear & ""
con.Execute "update SLEDGER set Owner=0"
DoEvents
DoEvents


Dim r1, r2 As Integer

r2 = 0
r1 = 1

If fillVs.State = 1 Then fillVs.close
fillVs.Open "select DISTCODE as City,SUBLEDGER as Party,op,drcr from SLEDGER where gledger='SUNDRY DEBTORS'  order by SUBLEDGER", con, adOpenDynamic, adLockOptimistic
If fillVs.EOF = False Then
vsop.rows = fillVs.RecordCount + 1
DoEvents
DoEvents
abc.Caption = vsop.rows
For I = 1 To vsop.rows - 1
  
  op = 0
  dr = 0
  CR = 0
  op = IIf(IsNull(fillVs(2)), 0, fillVs(2))
  
  
  
  If RS.State = 1 Then RS.close
  RS.Open "select sum(dr),sum(cr) from tempLedger2 where party='" & fillVs(1) & "'", con, adOpenDynamic, adLockOptimistic
  If Not IsNull(RS(0)) Then
     dr = RS(0)
     
  End If
  
  If Not IsNull(RS(1)) Then
     CR = RS(1)
  End If
  If fillVs(3) = "Cr" Then
    op = (-1 * fillVs(2))
  End If
  
  
  '-------------------------------------------------------
  
  amt = 0
  notmatch = False
  
  
  rs_Acc.MoveFirst
  rs_Acc.Find "subledger='" & fillVs!party & "'"
  If rs_Acc.EOF = False Then
     
     If rs_Acc!Opening < 0 Then
        amt = (rs_Acc!dramt - (Abs(rs_Acc!Opening) + rs_Acc!cramt))
     Else
        amt = ((rs_Acc!Opening + rs_Acc!dramt) - rs_Acc!cramt)
     End If
     
  Else
     amt = 0
  End If
   
   
  '-------------------------------------------------------
  If (Check3_filter.value = 1) Then
     
    
    
    a1 = Round((op + (dr - CR)))
    If (Round(amt, 0) <> Round((op + (dr - CR)))) Then
       r2 = r2 + 1
       notmatch = True
       If (r2 <= 1) Then
         vsop.rows = 2
       End If
    End If
  
     If (notmatch = True) Then
      vsop.rows = vsop.rows + 1
      X = Round(op + (dr - CR))
      vsop.TextMatrix(r1, 0) = fillVs(0) & ""
      vsop.TextMatrix(r1, 1) = fillVs(1)
      vsop.TextMatrix(r1, 2) = Format(Round(fillVs(2), 2), "0.00")
      vsop.TextMatrix(r1, 3) = fillVs(3) & ""
      vsop.TextMatrix(r1, 6) = Round(amt, 2)
    
      drcr = Round((op + (dr - CR)), 2)
      If Val(drcr) < 0 Then
         vsop.TextMatrix(r1, 4) = Abs(Round((op + (dr - CR)), 2))
         vsop.TextMatrix(r1, 5) = "Cr"
      Else
         vsop.TextMatrix(r1, 4) = Round((op + (dr - CR)), 2)
         vsop.TextMatrix(r1, 5) = "Dr"
      End If
      
      r1 = r1 + 1
      
      
     End If
  
  Else
  
  
      X = Round(op + (dr - CR))
      vsop.TextMatrix(I, 0) = fillVs(0) & ""
      vsop.TextMatrix(I, 1) = fillVs(1)
      vsop.TextMatrix(I, 2) = Format(Round(fillVs(2), 2), "0.00")
      vsop.TextMatrix(I, 3) = fillVs(3) & ""
      vsop.TextMatrix(I, 6) = Round(amt, 2)
    
      drcr = Round((op + (dr - CR)), 2)
      If Val(drcr) < 0 Then
         vsop.TextMatrix(I, 4) = Abs(Round((op + (dr - CR)), 2))
         vsop.TextMatrix(I, 5) = "Cr"
      Else
         vsop.TextMatrix(I, 4) = Round((op + (dr - CR)), 2)
         vsop.TextMatrix(I, 5) = "Dr"
      End If
  
  End If
  
  

  
 
    
  drcr = Format(Round(drcr, 2), "0.00")
  If Val(drcr) < 0 Then
  con.Execute "update sledger set Owner=" & Round(Val(drcr), 2) & " where subledger='" & fillVs(1) & "'"
  If bb = True Then
      contr.Execute "update sledger set OP=" & Abs(Round(Val(drcr), 2)) & ",drcr='Cr' where code='" & Trim(Mid(fillVs(1), 1, 6)) & "'"
  End If
  Else
      con.Execute "update sledger set Owner=" & Round(Val(drcr), 2) & " where subledger='" & fillVs(1) & "'"
  If bb = True Then
     contr.Execute "update sledger set OP=" & Abs(Round(Val(drcr), 2)) & ",drcr='Dr' where code='" & Trim(Mid(fillVs(1), 1, 6)) & "'"
  End If
  End If
  
  
  
  If Not IsNull(RS.Fields(1).value) Then
  If RS.Fields(1).value = 0 Then
  con.Execute "update sledger set Offdays='" & "1" & "' where subledger='" & fillVs(1) & "'"
  Else
  con.Execute "update sledger set Offdays='" & "2" & "' where subledger='" & fillVs(1) & "'"
  End If
  End If
  
  
  fillVs.MoveNext
  DoEvents
  DoEvents
  abc.Caption = abc.Caption - 1
  
  
  
Next

End If

vsop.Cols = 7
vsop.TextMatrix(0, 0) = "City"
vsop.TextMatrix(0, 1) = "Party"
vsop.TextMatrix(0, 2) = "Opening"
vsop.TextMatrix(0, 3) = "Dr/Cr"
vsop.TextMatrix(0, 4) = "Closing Balance"
vsop.TextMatrix(0, 5) = "Dr/Cr"
vsop.TextMatrix(0, 6) = "Acc Balance"

vsop.ColWidth(0) = 1800
vsop.ColWidth(1) = 3200
vsop.ColWidth(2) = 1200
vsop.ColWidth(3) = 500
vsop.ColWidth(4) = 1200
vsop.ColWidth(5) = 500
vsop.ColWidth(6) = 1200
abc.Caption = ""
bb = False
Screen.MousePointer = vbDefault

Exit Sub

aa11:
MsgBox "" & "Connection Not Created Properly !", vbInformation
Screen.MousePointer = vbDefault

End Sub
Sub showDataAsOn(D As Date)

Dim contr As New ADODB.Connection
Dim fillVs As New ADODB.Recordset
Dim dr, CR As Double
Dim bb As Boolean
Dim notmatch As Boolean

Screen.MousePointer = vbHourglass

'On Error GoTo aa11

If Me.txtpath.text <> "" Then
If MsgBox("Want To Transfer Closing ?", vbQuestion + vbYesNo) = vbYes Then
   contr.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source =" & txtpath.text
   contr.CursorLocation = adUseClient
   contr.Open
   bb = True
   contr.Execute "update sledger set OP=" & 0 & ",drcr=''"
End If
End If




If MsgBox("Want To Show Balance", vbInformation + vbYesNo) <> vbYes Then
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Dim op, drcr
Dim rs1 As New ADODB.Recordset



Dim dt1, dt2 As String
Dim rs_Acc As New ADODB.Recordset

dt1 = Format(fromdate.value, "MM/dd/yyyy")
dt2 = Format(dateAson.value, "MM/dd/yyyy")

Set rs_Acc = New ADODB.Recordset
Set rs_Acc = New ADODB.Recordset
Set rs_Acc = con.Execute("exec spBalanceSheet_Mearging_ch '" & dt1 & "','" & dt2 & "','SUNDRY DEBTORS'")



con.Execute "delete from templedger2 where " & stringyear & ""
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party,setupid,fyear)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER," & setupid & ",'" & session & "' from INVOICEA WHERE " & stringyear & " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & dateAson.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party,setupid,fyear) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,SUBLEDGER," & setupid & ",'" & session & "' from CREDITA WHERE " & stringyear & " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & dateAson.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party,setupid,fyear) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER," & setupid & ",'" & session & "' from CASHA where " & stringyear & " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & dateAson.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party,setupid,fyear) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD," & setupid & ",'" & session & "' from dnfa where " & stringyear & " and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & dateAson.value & "',103)"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,cr,dr,Party,setupid,fyear) Select cnd,'CN',cnn,'Credit Note',na,'0',psld," & setupid & ",'" & session & "' from Cnf1a where " & stringyear & " and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & dateAson.value & "',103)"
con.Execute "INSERT INTO tempLedger2 (dates,Billtype,bill,des,dr,cr,Party,setupid,fyear) Select dates,'J',Recno,Particullar,Dr,CR,PartyName," & setupid & ",'" & session & "' from ReceiveIssueParty where " & stringyear & " and convert(smalldatetime,dates,103)<=convert(smalldatetime,'" & dateAson.value & "',103) and firm='" & firm & "'"

DoEvents
DoEvents
con.Execute "update SLEDGER set Owner=0 where " & stringyear & ""
DoEvents
DoEvents


Dim r1, r2 As Integer

r2 = 0
r1 = 1


If fillVs.State = 1 Then fillVs.close
fillVs.Open "select DISTCODE as City,SUBLEDGER as Party,op,drcr from SLEDGER where " & stringyear & " and gledger='SUNDRY DEBTORS'", con, adOpenDynamic, adLockOptimistic
If fillVs.EOF = False Then
vsop.rows = fillVs.RecordCount
DoEvents
DoEvents
abc.Caption = vsop.rows
For I = 1 To vsop.rows - 1
  


  op = 0
  dr = 0
  CR = 0
  op = IIf(IsNull(fillVs(2)), 0, fillVs(2))
  
  
  If RS.State = 1 Then RS.close
  RS.Open "select sum(dr),sum(cr) from tempLedger2 where " & stringyear & " and party='" & fillVs(1) & "'", con, adOpenDynamic, adLockOptimistic
  If Not IsNull(RS(0)) Then
     dr = RS(0)
  End If
  
  If Not IsNull(RS(1)) Then
     CR = RS(1)
  End If
  If fillVs(3) = "Cr" Then
    op = (-1 * fillVs(2))
  End If
  
  
 '-------------------------------------------------------
  
  amt = 0
  notmatch = False
  
  
  rs_Acc.MoveFirst
  rs_Acc.Find "subledger='" & fillVs!party & "'"
  If rs_Acc.EOF = False Then
     
     If rs_Acc!Opening < 0 Then
        amt = (rs_Acc!dramt - (Abs(rs_Acc!Opening) + rs_Acc!cramt))
     Else
        amt = ((rs_Acc!Opening + rs_Acc!dramt) - rs_Acc!cramt)
     End If
     
  Else
     amt = 0
  End If
   
   
  '-------------------------------------------------------
  
  
   If (Check3_filter.value = 1) Then
     
    
    
    a1 = Round((op + (dr - CR)))
    
    If (Round(amt, 0) <> Round((op + (dr - CR)))) Then
    
         r2 = r2 + 1
         notmatch = True
         If (r2 <= 1) Then
           vsop.rows = 2
         End If
       
    End If
    
    
  
     If (notmatch = True) Then
      
          vsop.rows = vsop.rows + 1
          X = Round(op + (dr - CR))
          vsop.TextMatrix(r1, 0) = fillVs(0) & ""
          vsop.TextMatrix(r1, 1) = fillVs(1)
          vsop.TextMatrix(r1, 2) = Format(Round(fillVs(2), 2), "0.00")
          vsop.TextMatrix(r1, 3) = fillVs(3) & ""
          vsop.TextMatrix(r1, 6) = Round(amt, 2)
        
          drcr = Round((op + (dr - CR)), 2)
          If Val(drcr) < 0 Then
             vsop.TextMatrix(r1, 4) = Abs(Round((op + (dr - CR)), 2))
             vsop.TextMatrix(r1, 5) = "Cr"
          Else
             vsop.TextMatrix(r1, 4) = Round((op + (dr - CR)), 2)
             vsop.TextMatrix(r1, 5) = "Dr"
          End If
          
          r1 = r1 + 1
      
      
     End If
  
  Else
  
  
      X = Round(op + (dr - CR))
      vsop.TextMatrix(I, 0) = fillVs(0) & ""
      vsop.TextMatrix(I, 1) = fillVs(1)
      vsop.TextMatrix(I, 2) = Format(Round(fillVs(2), 2), "0.00")
      vsop.TextMatrix(I, 3) = fillVs(3) & ""
      vsop.TextMatrix(I, 6) = Round(amt, 2)
    
      drcr = Round((op + (dr - CR)), 2)
      If Val(drcr) < 0 Then
         vsop.TextMatrix(I, 4) = Abs(Round((op + (dr - CR)), 2))
         vsop.TextMatrix(I, 5) = "Cr"
      Else
         vsop.TextMatrix(I, 4) = Round((op + (dr - CR)), 2)
         vsop.TextMatrix(I, 5) = "Dr"
      End If
  
  End If
 
  
  
  'drcr = Round((op + (dr - cr)), 2)
  'If Val(drcr) < 0 Then
  '   vsop.TextMatrix(I, 4) = Abs(Round((op + (dr - cr)), 2))
  '   vsop.TextMatrix(I, 5) = "Cr"
  'Else
  '   vsop.TextMatrix(I, 4) = Round((op + (dr - cr)), 2)
  '   vsop.TextMatrix(I, 5) = "Dr"
  'End If
  
  
  drcr = Format(Round(drcr, 2), "0.00")
  If Val(drcr) < 0 Then
  con.Execute "update sledger set Owner=" & Round(Val(drcr), 2) & " where subledger='" & fillVs(1) & "'"
  If bb = True Then
      contr.Execute "update sledger set OP=" & Abs(Round(Val(drcr), 2)) & ",drcr='Cr' where code='" & Trim(Mid(fillVs(1), 1, 6)) & "'"
  End If
  Else
      con.Execute "update sledger set Owner=" & Round(Val(drcr), 2) & " where subledger='" & fillVs(1) & "'"
  If bb = True Then
     contr.Execute "update sledger set OP=" & Abs(Round(Val(drcr), 2)) & ",drcr='Dr' where code='" & Trim(Mid(fillVs(1), 1, 6)) & "'"
  End If
  End If
  
  
  
  If Not IsNull(RS.Fields(1).value) Then
  If RS.Fields(1).value = 0 Then
  con.Execute "update sledger set Offdays='" & "1" & "' where subledger='" & fillVs(1) & "'"
  Else
  con.Execute "update sledger set Offdays='" & "2" & "' where subledger='" & fillVs(1) & "'"
  End If
  End If
  
  
  fillVs.MoveNext
  DoEvents
  DoEvents
  abc.Caption = abc.Caption - 1
  
Next
End If

vsop.Cols = 7
vsop.TextMatrix(0, 0) = "City"
vsop.TextMatrix(0, 1) = "Party"
vsop.TextMatrix(0, 2) = "Opening"
vsop.TextMatrix(0, 3) = "Dr/Cr"
vsop.TextMatrix(0, 4) = "Closing Balance"
vsop.TextMatrix(0, 5) = "Dr/Cr"
vsop.TextMatrix(0, 6) = "A/c Balance"

vsop.ColWidth(0) = 1800
vsop.ColWidth(1) = 3200
vsop.ColWidth(2) = 1200
vsop.ColWidth(3) = 500
vsop.ColWidth(4) = 1200
vsop.ColWidth(5) = 500
vsop.ColWidth(6) = 1200

abc.Caption = ""
bb = False
Screen.MousePointer = vbDefault

'Exit Sub

'aa11:
'MsgBox "" & "Connection Not Created Properly !", vbInformation
'Screen.MousePointer = vbDefault

End Sub

Private Sub cmdShowClosing_Click()


bb1 = False
showData
Command2.Enabled = True
End Sub

Private Sub cmdTTrans_Click()
'frmTotalTrans.Show
End Sub

Private Sub cmdTitleLedger_Click()
frmTitleLedger.Show
End Sub

Private Sub cmdupdatep_Click()
   
   Dim partyname
   Dim pcode
   partyname = ""
   pcode = ""
   
    
   If RS.State = 1 Then RS.close
   RS.Open "select subledger from sledger where " & stringyear & "", con
   While RS.EOF = False
       
       aa = InStr(RS(0), " ")
       partyname = Mid(RS(0), aa)
       pcode = Mid(RS(0), 1, aa)
       
       con.Execute "update  Sledger  set party='" & Trim(partyname) & "',code='" & Trim(pcode) & "' where " & stringyear & " and subledger='" & RS(0) & "'"
       
       RS.MoveNext
       
   Wend
   
End Sub

Private Sub cmdUpDatePromotion_Click()

''''''On Error GoTo aa_
''''''
''''''Dim scid_ As String
''''''Dim donid_ As String
''''''
''''''Screen.MousePointer = vbHourglass
''''''For k_1 = 1 To vs.Rows - 1
''''''    donid_ = vs.TextMatrix(k_1, 1)
''''''    scid_ = Mid(vs.TextMatrix(k_1, 3), InStr(vs.TextMatrix(k_1, 3), ":") + 1)
''''''    updateDonnation scid_, donid_, Str(k_1)
''''''Next
''''''Screen.MousePointer = Default
''''''
''''''Exit Sub
''''''aa_:
''''''Screen.MousePointer = Default
''''''MsgBox "" & err.DESCRIPTION

    If Option2_donation.value = True Then
      vs.Clear
      Call cmdshow_Click
    End If
   
End Sub

Private Sub Command1_Click()
   
  If RS.State = 1 Then RS.close
  RS.Open "select * from pass where pass='" & cp & "'", con
  If RS.EOF = True Then
     MsgBox "Enter Valid Password !!", vbInformation
     Exit Sub
  
  Else

   Screen.MousePointer = vbHourglass
   
   On Error Resume Next
   
   For I = 1 To vsop.rows - 1
       If vsop.TextMatrix(I, 1) <> "" Then
          con.Execute "update SLEDGER set op=" & CDbl(vsop.TextMatrix(I, 2)) & ",drcr='" & vsop.TextMatrix(I, 3) & "' where SUBLEDGER='" & vsop.TextMatrix(I, 1) & "'"
       End If
   Next
   
   Screen.MousePointer = vbDefault
   

   
End If
   

   
   
   
   
   
End Sub



Private Sub Command2_Click()



Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String

If Not IsDate(date2) Then
   MsgBox "Plz. Enter Date...", vbCritical
   date2.SetFocus
   Exit Sub
End If



On Error GoTo err:



Screen.MousePointer = vbHourglass


If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double
Dim b1 As Boolean
Dim rs_ As New ADODB.Recordset

b1 = False
c = 1
r = 1


If rs_.State = 1 Then rs_.close
rs_.Open "select subledger,distcode,states  from SLEDGER where gledger='SUNDRY DEBTORS'", con, adOpenDynamic, adLockOptimistic

con.Execute "exec tmpdataForSale "


row_ = 1
col_ = 1

 xl.Columns("A:H").ColumnWidth = 12
 J = 2
 
 xlSheet.Cells(row_, 1).value = "Code"
 xlSheet.Cells(row_, 2).value = "Party Name"
 xlSheet.Cells(row_, 3).value = "Area"
 xlSheet.Cells(row_, 4).value = "State"
 xlSheet.Cells(row_, 5).value = "Closing Amt"
 xlSheet.Cells(row_, 6).value = "DR/CR"
 
 
 row_ = row_ + 1
 
 For I = 1 To vsop.rows - 1
           
    

     If rs1.State = 1 Then rs1.close
     rs1.Open "select distinct Party from tmpINVB_CrB where (convert(smalldatetime,INVOICEDATE,103)>convert(smalldatetime,'" & date2.text & "' ,103) and party ='" & vsop.TextMatrix(I, 1) & "' and Status_='I')", con
     If rs1.EOF = True Then
      
      If Val(vsop.TextMatrix(I, 4)) > 0 Then
         If vsop.TextMatrix(I, 5) = "Dr" Then
              
         
         
               xlSheet.Cells(row_, 1).value = Trim(Mid(vsop.TextMatrix(I, 1), 1, 6))
               xlSheet.Cells(row_, 2).value = Trim(Mid(vsop.TextMatrix(I, 1), 6))
               
               rs_.MoveFirst
               rs_.Find "subledger='" & vsop.TextMatrix(I, 1) & "'"
               If rs_.EOF = False Then
               
               xlSheet.Cells(row_, 3).value = rs_!distcode
               xlSheet.Cells(row_, 4).value = rs_!states
               
               End If
               
               xlSheet.Cells(row_, 5).value = vsop.TextMatrix(I, 4)
               xlSheet.Cells(row_, 6).value = vsop.TextMatrix(I, 5)
               
               row_ = row_ + 1
           End If
       End If
       
    End If
            
     
 Next
    


Screen.MousePointer = vbDefault


Exit Sub
Screen.MousePointer = vbDefault
err:
MsgBox err.DESCRIPTION



'''crpt.Reset
'''crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseDrClosing.rpt"
'''crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"
'''crpt.ReplaceSelectionFormula "{tempLedgerRpt.Offdays}='" & "1" & "' and {tempLedgerRpt.Owner}>=" & 1 & ""
'''DoEvents
'''MsgBox ("View")
'''crpt.Formulas(0) = "partyname='" & cboStation1.Text & "'"
'''crpt.WindowShowPrintSetupBtn = True
'''crpt.WindowShowPrintBtn = True
'''crpt.WindowState = crptMaximized
'''crpt.WindowShowSearchBtn = True
'''crpt.Action = 1

End Sub

Private Sub Command3_Click()
 
 If cboStation.text = "" Then
    MsgBox "Please Select Station...", vbInformation
    Exit Sub
 End If
 
 Screen.MousePointer = vbHourglass
 'login.DSN
 CityWiseStatement
 cboPartyList.Visible = False
 Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()

Dim FSO As filesystemobject
Dim f As File
Dim txt As TextStream
Dim matter As String
Dim Total As String
Dim s(1, 2) As String
Set FSO = New filesystemobject
Dim ss As String
'
Dim s1

matter = ""

Set txt = FSO.CreateTextFile(App.Path & "\mobile.txt", True)

If RS.State = 1 Then RS.close
If Check2.value = 0 Then
RS.Open "select mobile from sledger where " & stringyear & " and distcode='" & cboStation1.text & "'", con, adOpenKeyset, adLockReadOnly
Else
RS.Open "select mobile from sledger where " & stringyear & " and states='" & cboStation1.text & "'", con, adOpenKeyset, adLockReadOnly
End If

While RS.EOF = False


If Len(RS(0)) > 0 Then

s1 = Split(RS(0), ",")
For I = 0 To UBound(s1)
    matter = matter & Trim(s1(I)) & vbNewLine
Next



End If
RS.MoveNext
Wend

txt.Write matter
txt.close

MsgBox "File Created ....", vbInformation

Shell App.Path & "\notepad.exe " & App.Path & "\mobile.txt", vbMaximizedFocus

End Sub

Private Sub Command5_Click()



'''''

Dim P1 As Integer

P1 = Right(session, 2)

If P1 >= 23 Then

VS_sale.Visible = True

Dim Salefrom, SaleTo, ReturnFrom, ReturnTo As String
Dim pcode, database_last_ As String

Dim db1

pcode = Left(cboParty.text, 5)
If RS.State = 1 Then RS.close
RS.Open "select fromDate,toDate,fromDateSRet,toDateSRet,NotCreated,Current_Next,DataBase from turnOverDis order by Current_Next", CCON
If RS.EOF = False Then


Salefrom = RS!fromdate
RS.MoveNext
SaleTo = RS!todate
ReturnFrom = RS!fromDateSRet
ReturnTo = RS!toDateSRet
db1 = RS!NotCreated

If db1 = "n" Then
   next_dbase = "n"
End If


saleAmt1 = 0
RetAmt1 = 0
Set rs1 = New ADODB.Recordset
Set rs1 = con.Execute("exec Sp_tmpSaleRegister_ForLedger '" & pcode & "','" & next_dbase & "','I','" & Format(Salefrom, "mm/dd/yyyy") & "','" & Format(SaleTo, "mm/dd/yyyy") & "','" & Format(fromdate, "mm/dd/yyyy") & "','" & Format(todate, "mm/dd/yyyy") & "'")
If rs1.EOF = False Then
   saleAmt1 = Round(rs1(1))
End If



Set rs1 = New ADODB.Recordset
Set rs1 = con.Execute("exec Sp_tmpSaleRegister_ForLedger '" & pcode & "','" & next_dbase & "','C','" & Format(Salefrom, "mm/dd/yyyy") & "','" & Format(SaleTo, "mm/dd/yyyy") & "','" & Format(ReturnFrom, "mm/dd/yyyy") & "','" & Format(ReturnTo, "mm/dd/yyyy") & "'")
If rs1.EOF = False Then
   RetAmt1 = Round(rs1(1))
End If


VS_sale.TextMatrix(0, 1) = saleAmt1
VS_sale.TextMatrix(1, 1) = RetAmt1
VS_sale.TextMatrix(2, 1) = (saleAmt1 - RetAmt1)



End If




''''''
Else

VS_sale.Visible = False

End If



seriesWiseDiscount

If Mid(session, 6) >= 18 Then
    Screen.MousePointer = vbHourglass
    
    If InStr(cboParty.text, "(EM)") > 0 Then
       updateGP
    End If
    
    PartyLedgerNew
    CalculateTotalDrCrNew
    txtdes.SetFocus
    Screen.MousePointer = vbDefault
Else
    MsgBox "Ledger View(New & fast) " & vbCrLf & "Option is not working in this session..."
    Check3_newledger.value = 0
End If



End Sub
Sub updateGP()
    
Dim s_ As String
s_ = ""

If RS.State = 1 Then RS.close
RS.Open "SELECT INVOICENO FROM invoicea where SUBLEDGER='" & cboParty.text & "' order by INVOICENO", con
While RS.EOF = False

    s_ = ""
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT distinct GROUPCODE FROM invoiceBQry where INVOICENO='" & RS!invoiceNo & "'", con
    While rs1.EOF = False
    
      If s_ = "" Then
         s_ = rs1!groupcode
      Else
         s_ = s_ & "," & rs1!groupcode
      End If
    
    rs1.MoveNext
    Wend
    
    con.Execute "update INVOICEA set gpName='" & s_ & "' where INVOICENO='" & RS!invoiceNo & "'"

 RS.MoveNext
Wend
 
 
 
s_ = ""

If RS.State = 1 Then RS.close
RS.Open "SELECT INVOICENO FROM credita where SUBLEDGER='" & cboParty.text & "' order by INVOICENO", con
While RS.EOF = False

    s_ = ""
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT distinct GROUPCODE FROM CREDITBQry where INVOICENO='" & RS!invoiceNo & "'", con
    While rs1.EOF = False
    
      If s_ = "" Then
         s_ = rs1!groupcode
      Else
         s_ = s_ & "," & rs1!groupcode
      End If
    
    rs1.MoveNext
    Wend
    
    con.Execute "update credita set gpName='" & s_ & "' where INVOICENO='" & RS!invoiceNo & "'"

 RS.MoveNext
Wend
  
    
End Sub
Private Sub Command5_print_Click()

con.Execute "update DonnationMain set Print_='' where Print_='Print'"


For k1 = 1 To vs_promotion.rows - 1
   If vs_promotion.TextMatrix(k1, 0) = "Print" Then
      con.Execute "update DonnationMain set Print_='Print',remarks1='" & vs_promotion.TextMatrix(k1, 9) & "' where DNo=" & vs_promotion.TextMatrix(k1, 2) & ""
   End If
Next


DoEvents
DSNNew

MsgBox ("Print")



crpt.Reset
crpt.ReportFileName = App.Path & "\reports\SpVoucher.rpt"
crpt.Connect = constr
crpt.ReplaceSelectionFormula "{DonnationMain.print_}='" & "Print" & "'"
crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.WindowShowSearchBtn = True
crpt.Action = 1


End Sub

Private Sub Command6_Click()

Dim strProgramName As String
Dim strArgument As String


con.Execute "delete from createpdf"
If cboParty.text <> "" Then

Code = Trim(Mid(cboParty.text, 1, 6))
party_ = Trim(Mid(cboParty.text, 6))

con.Execute "insert into createpdf(pname) values('" & party_ & ":" & Code & "')"
End If

DoEvents
DoEvents

If (session = "2021-22") Then

strProgramName = "\\192.168.0.140\blueprintSales\pdf_create_2122\bin\Debug\MailSystem.exe"
strArgument = "/G"

ElseIf (session = "2022-23") Then

strProgramName = "\\192.168.0.140\blueprintSales\pdf_create_2223\bin\Debug\MailSystem.exe"
strArgument = "/G"

ElseIf (session = "2023-24") Then

strProgramName = "\\192.168.0.140\blueprintSales\pdf_create_2324\bin\Debug\MailSystem.exe"
strArgument = "/G"

ElseIf (session = "2024-25") Then

strProgramName = "\\192.168.0.140\blueprintSales\pdf_create_2425\bin\Debug\MailSystem.exe"
strArgument = "/G"

End If



'Call Shell("""" & strProgramName & """ """ & strArgument & """", vbNormalFocus)
Call Shell(strProgramName, vbNormalFocus)








End Sub

Private Sub Command7_Click()
Screen.MousePointer = vbHourglass

Dim strProgramName As String
Dim strArgument As String


DoEvents
DoEvents


If (session = "2022-23") Then

strProgramName = "\\192.168.0.140\blueprintSales\PartyDocument\bin\Debug\MailSystem.exe"
strArgument = "/G"

End If

Call Shell(strProgramName, vbNormalFocus)






Screen.MousePointer = vbDefault

End Sub

Private Sub Command8print_Click()
'On Error GoTo aa10

Screen.MousePointer = vbHourglass
Dim op, drcr
Dim rs1 As New ADODB.Recordset
Dim rs1_rpt As New ADODB.Recordset
Dim rss_ds As New ADODB.Recordset
Dim inv_str As String

'login.DSN


con.Execute "delete from templedger1 where userid='" & UId & "'"

If rss_ds.State = 1 Then rss_ds.close
rss_ds.Open "select amount,invoiceno from INVOICEC where TEXT='SCHEME DISCOUNT' and AMOUNT>0", con


If rs1_rpt.State = 1 Then rs1_rpt.close
rs1_rpt.Open "select max(rptid) from tempLedger1", con, adOpenDynamic, adLockOptimistic
If IsNull(rs1_rpt(0)) Then
rptid = 9999
Else
rptid = rs1_rpt(0) + 1
End If

If RS.State = 1 Then RS.close
RS.Open "select subledger from SLEDGER where " & stringyear & " and subledger = '" + Trim(cboParty.text) + "'", con
While RS.EOF = False

'==Code For Opening=============================================

If lblCr = "dr" Then

    con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype,setupid,fyear,rptid,UserId)  SELECT op,drcr,subledger,'Opening','" & setupid & "','" & session & "'," & rptid & "," & UId & " from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'   group by op,subledger,drcr HAVING  op <> 0;"
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT op,drcr from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'", con
    If Not IsNull(rs1.Fields(0).value) Then
       op = Val(rs1.Fields(0).value)
       drcr = rs1.Fields(1).value
    Else
       op = 0
    End If

Else

 
    
    con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype,Bill,fyear,setupid,userid)  values(" & Val(txtOp) & ",'" & cboop & "','" & RS.Fields(0).value & "','Opening',0,'" & session & "','" & setupid & "','" & userid & "')"
    op = Val(txtOp)
    drcr = cboop
    
End If


'==============================================


con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER,Fyear,setupid," & rptid & ",'" & UId & "' from INVOICEA where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',BAA,netamount,SUBLEDGER,Fyear,setupid," & rptid & ",'" & UId & "' from CREDITA where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER,Fyear,setupid," & rptid & ",'" & UId & "' from CASHA where  " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD,Fyear,setupid," & rptid & ",'" & UId & "' from dnfa where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party,fyear,setupid,rptid,userid) Select cnd,'CN',cnn,'Credit Note',na,'0',psld,Fyear,setupid," & rptid & ",'" & UId & "' from Cnf1a where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select dates,'J',Recno,Particullar,Dr,CR,PartyName,Fyear,setupid," & rptid & ",'" & UId & "' from ReceiveIssueParty where " & stringyear & " and PartyName='" & RS.Fields(0).value & "' and firm='" & firm & "' order by dates,recno"

If lblCr = "cr" Then

con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select VoucherDate,VoucherType,VoucherNumber,DESCRIPTION,Amount,0,SubLedger,Fyear,setupid," & rptid & ",'" & UId & "' from vouchers where (" & stringyear & " and SubLedger='" & RS.Fields(0).value & "' and DebitorCredit='D') order by VoucherDate,VoucherNumber"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,rptid,userid) Select VoucherDate,VoucherType,VoucherNumber,DESCRIPTION,0,Amount,SubLedger,Fyear,setupid," & rptid & ",'" & UId & "' from vouchers where (" & stringyear & " and SubLedger='" & RS.Fields(0).value & "' and DebitorCredit='C') order by VoucherDate,VoucherNumber"

End If

'===============================================================
If op <> 0 Then
con.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "',UserId='" & UId & "' where " & stringyear & " and party = '" + RS.Fields(0).value + "' and Billtype<>'Opening'"
End If
'===============================================================
Sleep (200)
RS.MoveNext
Wend

'convert(smalldatetime,dates,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)

If rs1.State = 1 Then rs1.close
rs1.Open "select gledger from SLEDGER where subledger='" & cboParty.text & "'", con
If rs1.EOF = False Then
If rs1!gledger = "SUNDRY DEBTORS" Then
    For J = 1 To vs1.rows - 1
    If vs1.TextMatrix(J, 2) <> "" Then
    
    
       s_1 = InStr(vs1.TextMatrix(J, 3), "(")
       s_2 = InStr(vs1.TextMatrix(J, 3), ")")
       con.Execute "Update templedger1 set des='" & vs1.TextMatrix(J, 3) & "' where (bill='" + vs1.TextMatrix(J, 1) + "' and Billtype='" + vs1.TextMatrix(J, 0) + "')"
       
       'End If
    End If
    Next
End If
End If

Sleep (300)


DSNNew




    'CommonDialog1.Flags = 64
    'CommonDialog1.ShowPrinter
 
    crpt.Reset
    crpt.ReportFileName = App.Path & "\reports\PartyLedger.rpt"
    
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.UserId}='" & UId & "'"
    
    
   
    crpt.Connect = constr
    crpt.WindowShowPrintSetupBtn = True
    crpt.WindowShowPrintBtn = True
    crpt.WindowState = crptMaximized
    'crpt.Destination = crptToPrinter
    crpt.Action = 1
    Screen.MousePointer = vbDefault

'Else
'
'   Screen.MousePointer = vbDefault
'   PopUpValue6 = cboParty.Text
'   popupvalue5 = rptid
'   popupvalue4 = "PartyLedger.rpt"
'   frmSendMail.Show 1
'
'End If




Exit Sub
aa10:
Screen.MousePointer = vbDefault
'MsgBox err.DESCRIPTION

End Sub

Private Sub crdit_Click()
    If crdit.value = True Then
       Call cmdshow_Click
    End If
End Sub

Private Sub credit_Click()
    If credit.value = True Then
       Call cmdshow_Click
    End If

End Sub

Private Sub dbit_Click()
   If dbit.value = True Then
       Call cmdshow_Click
    End If
End Sub
Sub fillDocument(code_ As String)

Dim DocDB As String
Dim lastYrs As String
 

lblCAF(1).Caption = ""
vs_document.Clear
vs_document.rows = 1


VS_sale.Clear
VS_sale.rows = 3

VS_sale.ColWidth(0) = 1100
VS_sale.ColWidth(1) = 850

VS_sale.TextMatrix(0, 0) = "T.Sale"
VS_sale.TextMatrix(1, 0) = "T.SReturn"
VS_sale.TextMatrix(2, 0) = "Net Sale"


lastYrs = ""

If (DateDiff("d", Now, SessionLastDate) <= 0) Then
   lastYrs = "current"
Else
   lastYrs = "last"
End If



If (lastYrs = "last") Then

    dk = Mid(session_last1, 6)
    If Val(dk) >= 24 Then
     caf_ = "MOU-" & session_last1
    Else
     caf_ = "MOU"
    End If

Else

    dk = Mid(session, 6)
    If Val(dk) >= 24 Then
     caf_ = "MOU-" & session
    Else
     caf_ = "MOU"
    End If

End If




k5 = 1
If rs1.State = 1 Then rs1.close

Set rs1 = New ADODB.Recordset
Set rs1 = con_doc1.Execute("exec SearchPDocument '" & code_ & "','" & caf_ & "'")


While rs1.EOF = False

vs_document.rows = vs_document.rows + 1
vs_document.TextMatrix(k5, 0) = rs1!linkname

If IsNull(rs1!fname) Then
  vs_document.TextMatrix(k5, 1) = "No"
 
Else
 vs_document.TextMatrix(k5, 1) = "Yes"
End If


k5 = k5 + 1


rs1.MoveNext
Wend


Set rs1 = New ADODB.Recordset
rs1.Open "SELECT fname FROM PartyDocument where code='" & code_ & "' and LinkName ='" & caf_ & "'", con_doc1
If rs1.EOF = True Then
   lblCAF(1).Caption = "MOU Not Uploaded.."
Else

   If IsNull(rs1!fname) Then
      lblCAF(1).Caption = "MOU Not Uploaded.."
   End If
   
End If

vs_document.FormatString = "Doc.Type|Upload"

vs_document.ColWidth(0) = 1150
vs_document.ColWidth(1) = 750



End Sub

Private Sub Form_Activate()
'login.DSN
'Me.WindowState = 2

Me.dataTrans.Visible = False
'Check3_ledgerClosingTrans.Visible = True
If main.UserName = "v" Then
'Me.dataTrans.Visible = True
Check3_ledgerClosingTrans.Visible = True
End If
   
Dim rs_1 As New ADODB.Recordset
If rs_1.State = 1 Then rs_1.close
rs_1.Open "select * from pass where pass='" & strledger & "'", con
If rs_1.EOF = False Then
   'txtRem.Visible = False
   cmdShow1.Visible = False
   txtrem.Visible = True
   txtrem.Enabled = True
Else
   txtrem.Visible = True
   txtrem.Enabled = False
   'cmdShow1.Visible = True
End If


If rs_1.State = 1 Then rs_1.close
rs_1.Open "select yarto from setup1", con
If rs_1.EOF = False Then
SessionLastDate1 = rs_1(0)
End If

   
If RS.State = 1 Then RS.close
RS.Open "select fromDate from turnOverDis", CCON
If RS.EOF = False Then
   SessionLastDate = RS(0)
End If
   
   
  
  
End Sub
Private Sub Form_Load()
ch_ = False
user_id = Trim((Sys_user_ + Str(UId)))

ch_din = "n"
PopUpValue3 = ""
PopUpValue1 = ""
PopUpValue2 = ""

Me.top = 20
Me.Left = 50

Me.Width = 17400
Me.Height = 10400


vsIni
On Error Resume Next

kk = 1
dateAson.value = Date

fromdate.value = Date
todate.value = Date
from_date = fromdate.value
'FillGrid

maxId
setWidth
cboop.ListIndex = 0

vsop.Cols = 7

vsop.TextMatrix(0, 0) = "City"
vsop.TextMatrix(0, 1) = "Party"
vsop.TextMatrix(0, 2) = "Opening"
vsop.TextMatrix(0, 3) = "Dr/Cr"
vsop.TextMatrix(0, 4) = "Closing Balance"
vsop.TextMatrix(0, 5) = "Dr/Cr"
vsop.TextMatrix(0, 6) = "A/c Balance"


'fillDocument ("")

vsop.ColWidth(0) = 1800
vsop.ColWidth(1) = 3200
vsop.ColWidth(2) = 1400
vsop.ColWidth(3) = 500
vsop.ColWidth(4) = 1400
vsop.ColWidth(5) = 500
vsop.ColWidth(6) = 1200

If RS.State = 1 Then RS.close
RS.Open "select yarfrom,yarto from setup1 where " & stringyear & "", con
If RS.EOF = False Then
   fromdate.value = RS.Fields(0).value
   If (DateValue(RS!yarfrom) <= DateValue(Date) And DateValue(RS!yarto) >= DateValue(Date)) Then
      RecDates.value = Date
   Else
      RecDates.value = RS.Fields(1).value
   End If
   txt_ason.value = RS!yarto
End If

Me.top = 50
Me.Left = 50

Opening.Tab = 1


If RS.State = 1 Then RS.close
RS.Open "select * from setup1 where " & stringyear & " ", con
If RS.EOF = False Then
    date1.text = RS!yarfrom
    date2.text = RS!yarto
End If

bb1 = False
fetchTab2

If module_ = "Invoicing" Then
    Me.Caption = "Chitra"
    Option_bookIssueSp.Visible = False
    Option_bookRetSp.Visible = False
    dbit.Visible = True
    crdit.Visible = True
    credit.Visible = True
    cash.Visible = True
    sales.Visible = True
    Option2_app.Visible = True
    Opening.TabEnabled(0) = True
    Opening.TabEnabled(1) = True
    Opening.TabEnabled(2) = True
    Option2_donation.Visible = True
    cmdRefProm.Visible = True
    cmdUpDatePromotion.Visible = True
Else
   Opening.Tab = 0
   Me.Caption = "Blue Print"
   Option_bookIssueSp.Visible = True
   Option_bookRetSp.Visible = True
   Option2_app.Visible = False
   dbit.Visible = False
   crdit.Visible = False
   credit.Visible = False
   cash.Visible = False
   sales.Visible = False
   Opening.TabEnabled(0) = True
   Opening.TabEnabled(1) = False
   Opening.TabEnabled(2) = False
   
   Option2_donation.Visible = False
   cmdRefProm.Visible = False
   cmdUpDatePromotion.Visible = False

   
End If

adddist

con.Execute "delete from tempLedger1 where UserId='" & UserName & "'"
con.Execute "delete from tmpDonnationnew where uid='" & user_id & "'"
con.Execute "delete from tmpSaladjust where uid='" & user_id & "'"
con.Execute "delete from templedger6 where userid='" & user_id & "'"


str_don = "n"
If rs_don.State = 1 Then rs_don.close
rs_don.Open "select LastDatabase from data", CCON
If rs_don.EOF = False Then
Set con_don = New ADODB.Connection
If LCase(server_) = "server" Then

   con_don.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs_don!LastDatabase & "; uid=" & sql_user & "; PWD=" & sql_pass
   'con.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & databaseNew & "; uid=" & sql_user & "; PWD=" & sql_pass

   con_don.Open
   str_don = "y"
Else
   con_don.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs_don!LastDatabase & "; UID=; PWD=;"
   con_don.Open
   str_don = "y"
End If
End If


If RS.State = 1 Then RS.close
RS.Open "select fromDate,toDate,NotCreated from turnOverDis where fyear='" & session & "'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
   fdate1_ = RS!fromdate
  End If
End If

genrateCloasing

sess1_ = Val(Right(session, 2))
If sess1_ >= 23 Then
con.Execute "exec Sp_UpdateCrItem_adj"
End If



If Mid(session, 6) >= 18 Then
   Check3_newledger.value = 1
Else
   Check3_newledger.value = 0
End If

Screen.MousePointer = vbDefault
BackColorFrom Me


FileDocument_Con

'Set con_doc1 = New ADODB.Connection
'DocDB = "Database=chitraData_2223"
'
'
'If LCase(server_) = "server" Then
'   con_doc1.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & DocDB & "; UID=" & sql_user & "; PWD=" & sql_pass
'End If
'
'DoEvents
'DoEvents
'
'
'con_doc1.CursorLocation = adUseClient
'If con_doc1.State = 1 Then con_doc1.close
'con_doc1.Open
'
'DoEvents
'DoEvents

End Sub
Sub setsecurity()

If LCase(strledger) <> "cp" Then
   cmdShow1.Visible = False
   MsgBox "Enter Valid Password !!", vbInformation
   Exit Sub
Else
   saveData
End If
   
End Sub
Private Sub Form_Resize()

panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2
'BackColorFrom Me, 1

'login.DSN

End Sub

Private Sub Label1_Click(Index As Integer)
If cboParty.text <> "" Then
    PopUpValue6 = cboParty.text
    
    If (LCase(UserName) = "admin") Then
    
        frmSeriesWiseDis.cmdSave_2.Enabled = True
        frmSeriesWiseDis.cmdDelete_3.Enabled = True
        frmSeriesWiseDis.cmdAdd_1.Enabled = True
        frmSeriesWiseDis.cmdEdit_4.Enabled = True
    Else
        frmSeriesWiseDis.cmdSave_2.Enabled = False
        frmSeriesWiseDis.cmdDelete_3.Enabled = False
        frmSeriesWiseDis.cmdAdd_1.Enabled = False
        frmSeriesWiseDis.cmdEdit_4.Enabled = False

    End If
    
    frmSeriesWiseDis.Show 1
Else
   MsgBox "Plz Search Party ..", vbInformation
End If
End Sub

Private Sub List1_ch_DblClick()

Dim ss1
ss1 = Split(List1_ch, ":")
cboParty.text = ss1(0)
List1_ch.Visible = False

End Sub

Private Sub Opening_Click(PreviousTab As Integer)
      
     ' Screen.MousePointer = vbHourglass
      Dim closing As Double
      
      
      closing = 0
      
      If Opening.Tab = 0 Then
         'frmBillList.Show
         Call cmdshow_Click
      ElseIf Opening.Tab = 1 Then
         ' BackColorFrom
         BackColorFrom Me
      End If
      
      
      
      'Screen.MousePointer = vbDefault
      
End Sub
Sub fetchTab2()

        Screen.MousePointer = vbHourglass

        Dim fillVs As New ADODB.Recordset
        If fillVs.State = 1 Then fillVs.close
        'fillvs.Open "select DISTCODE as City,SUBLEDGER as Party,op,drcr from closing where gledger='SUNDRY DEBTORS'", con
        fillVs.Open "SELECT SLEDGER.DISTCODE,SLEDGER.SUBLEDGER,SLEDGER.OP,SLEDGER.drcr,(Sum(templedger1.Dr)-Sum(templedger1.Cr)) AS bal1 FROM SLEDGER LEFT JOIN templedger1 ON SLEDGER.SUBLEDGER = templedger1.Party where gledger='SUNDRY DEBTORS' GROUP BY SLEDGER.SUBLEDGER,SLEDGER.DISTCODE,[SLEDGER.OP], SLEDGER.drcr, SLEDGER.gledger", con

        If fillVs.EOF = False Then
            vsop.rows = fillVs.RecordCount
            For I = 1 To vsop.rows - 1
              vsop.TextMatrix(I, 0) = fillVs(0) & ""
              vsop.TextMatrix(I, 1) = fillVs(1)
              vsop.TextMatrix(I, 2) = Format(fillVs(2), "0.00")
              vsop.TextMatrix(I, 3) = fillVs(3) & ""

              If Not IsNull(fillVs(4)) Then

                     If vsop.TextMatrix(I, 3) = "Cr" Then
                         vsop.TextMatrix(I, 4) = ((-1 * (vsop.TextMatrix(I, 2))) + fillVs(4))
                         If vsop.TextMatrix(I, 4) < 0 Then
                            vsop.TextMatrix(I, 4) = Format((-1 * Val(vsop.TextMatrix(I, 4))), "0.00")
                            vsop.TextMatrix(I, 5) = "Cr"
                         Else
                            vsop.TextMatrix(I, 5) = "Dr"
                            vsop.TextMatrix(I, 4) = Format(Val(vsop.TextMatrix(I, 4)), "0.00")
                         End If

                     Else
                         vsop.TextMatrix(I, 4) = ((Val(vsop.TextMatrix(I, 2))) + fillVs(4))
                         If vsop.TextMatrix(I, 4) < 0 Then
                            vsop.TextMatrix(I, 4) = Format((-1 * Val(vsop.TextMatrix(I, 4))), "0.00")
                            vsop.TextMatrix(I, 5) = "Cr"
                         Else
                            vsop.TextMatrix(I, 5) = "Dr"
                            vsop.TextMatrix(I, 4) = Format(Val(vsop.TextMatrix(I, 4)), "0.00")
                         End If


                     End If
              End If


              fillVs.MoveNext
            Next
        End If

        vsop.Cols = 6
        vsop.TextMatrix(0, 0) = "City"
        vsop.TextMatrix(0, 1) = "Party"
        vsop.TextMatrix(0, 2) = "Opening"
        vsop.TextMatrix(0, 3) = "Dr/Cr"
        vsop.TextMatrix(0, 4) = "Closing"
        vsop.TextMatrix(0, 5) = "Dr/Cr"


        vsop.ColWidth(0) = 1800
        vsop.ColWidth(1) = 3600
        vsop.ColWidth(2) = 1200
        vsop.ColWidth(3) = 500
        vsop.ColWidth(4) = 1200
        vsop.ColWidth(5) = 500

        Screen.MousePointer = vbDefault



End Sub

Private Sub Option_bookIssueSp_Click()
    If Option_bookIssueSp.value = True Then
       Call cmdshow_Click
    End If
End Sub

Private Sub Option_bookRetSp_Click()
    If Option_bookRetSp.value = True Then
       Call cmdshow_Click
    End If

End Sub

Private Sub Option1_Click()
   If Option1.value = True Then
      bill.Visible = True
   Else
      partydrcr.Visible = False
      bill.Visible = True
   End If
End Sub

Private Sub Option2_Click()
   If Option2.value = 1 Then
      txtadmin.Visible = True
      Label14.Visible = True
   Else
      txtadmin.Visible = False
      Label14.Visible = False
   End If
End Sub

Private Sub Option2_agm_Click()
   If Option2_agm.value = True Then
       Call cmdshow_Click
    End If
End Sub

Private Sub Option2_app_Click()
    If Option2_app.value = True Then
       Call cmdshow_Click
    End If

End Sub

Private Sub Option2_donation_Click()
    If Option2_donation.value = True Then
      vs.Clear
      Call cmdshow_Click
    End If
    
End Sub

Private Sub Option2_mzn_Click()
adddist
End Sub

Private Sub Option2_state_Click()
adddist
End Sub
Sub adddist()


cboStation.Clear

If Option3_dist.value = True Then

Command3.Caption = "&Ledger Dist. Wise"
Label15.Caption = Option3_dist.Caption

If RS.State = 1 Then RS.close
RS.Open "select distinct(DISTCODE) from SLEDGER where " & stringyear & " and DISTCODE<>''", con
While RS.EOF = False
   cboStation.AddItem RS.Fields(0).value
   cboStation1.AddItem RS.Fields(0).value
   RS.MoveNext
Wend

ElseIf Option2_state.value = True Then

Command3.Caption = "&Ledger State Wise"
Label15.Caption = Option2_state.Caption

If RS.State = 1 Then RS.close
RS.Open "select states from SLEDGER group by states", con
While RS.EOF = False
cboStation.AddItem RS(0)
RS.MoveNext
Wend
cboStation.AddItem "ALL"


ElseIf Option4_rep.value = True Then

Command3.Caption = "&Ledger Rep Wise"
Label15.Caption = Option4_rep.Caption

If RS.State = 1 Then RS.close
RS.Open "select Rep as Representative from SalesRepQry order by Rep", CON_blue

    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cboStation.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    
    Me.cboStation.AddItem "ALL"
   
ElseIf Option2_mzn.value = True Then

Command3.Caption = "&Ledger Rep Wise"
Label15.Caption = Option2_mzn.Caption

If RS.State = 1 Then RS.close
RS.Open "select Manager from rep where len(Manager)>0 group by Manager", CON_blue

    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cboStation.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    

    
End If

End Sub

Private Sub Option3_dist_Click()
adddist
End Sub

Private Sub Option4_rep_Click()
adddist
End Sub

Private Sub party_Click()
   
   If party.value = True Then
      bill.Visible = False
      frmReceiveFromParty.Show
   Else
      partydrcr.Visible = False
      bill.Visible = True
   End If
   
   frmReceiveFromParty.top = 800

End Sub
Private Sub RecDates_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        set_focus = False
        cboParty.SetFocus
     End If
End Sub
Private Sub SSTab1_DblClick()
   RecDates.SetFocus
End Sub

Private Sub RecDates_LostFocus()
    
    If Trim(RecDates.value) <> "" Then
        If Not checkdate(Trim(RecDates.value), RecDates) Then
            RecDates.SetFocus
        End If
    End If

End Sub

Private Sub sales_Click()
    If sales.value = True Then
       Call cmdshow_Click
    End If
End Sub

Private Sub selectAll_Click()
If selectAll.value = 1 Then
    For I = 0 To cboPartyList.ListCount - 1
        cboPartyList.Selected(I) = True
    Next
Else
   For I = 0 To cboPartyList.ListCount - 1
    cboPartyList.Selected(I) = False
   Next
End If
End Sub
Private Sub Timer1_Timer()
Static L As Integer

If L = 0 Then
    Label1(20).ForeColor = vbYellow
    Label1(21).ForeColor = vbBlue
    L = 1
    Exit Sub
ElseIf L = 1 Then
    Label1(20).ForeColor = vbBlue
    Label1(21).ForeColor = vbYellow
    L = 0
    Exit Sub
End If


End Sub

Private Sub Timer2_Timer()
Static L As Integer

If L = 0 Then
    Label1(1).ForeColor = vbYellow
    Label1(2).ForeColor = vbBlue
    L = 1
    Exit Sub
ElseIf L = 1 Then
    Label1(1).ForeColor = vbBlue
    Label1(2).ForeColor = vbYellow
    L = 0
    Exit Sub
End If

End Sub

Private Sub txtadmin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   setsecurity
   pass.Visible = False
End If
End Sub



Private Sub txtchequeeNo_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 13 Then
  List1_ch.Clear
  If rs1.State = 1 Then rs1.close
  rs1.Open "SELECT PartyName,cr,dr FROM ReceiveIssueParty where Particullar like '%" & txtchequeeNo.text & "%'", con
  While rs1.EOF = False
     cboParty.text = rs1(0)
    
     If rs1!CR > 0 Then
        List1_ch.AddItem rs1(0) & ":" & "Cr"
     End If
     
     If rs1!dr > 0 Then
        List1_ch.AddItem rs1(0) & ":" & "Dr"
     End If

     
     rs1.MoveNext
  Wend
  
  List1_ch.Visible = True
  
End If


End Sub

Private Sub txtdes_GotFocus()
  txtdes.BackColor = &HFFFFC0
End Sub

Private Sub txtdes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtQty.SetFocus
  End If
End Sub

Private Sub txtEnterPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Option2_donation.value = True Then
    'strName = InputBox("Enter password for change sponsorship status :", "Input Value")
    strName = encrypt(txtEnterPass)
    If RS.State = 1 Then RS.close
    RS.Open "select * from pass where pass='" & cp & "' and donnation='" & strName & "'", con
    If RS.EOF = False Then
        ch_din = "y"
    Else
        ch_din = "n"
    End If
    frmPassword.Visible = False
End If

   
End If
End Sub

Private Sub txtOp_GotFocus()
txtOp.BackColor = &HFFFFC0
End Sub
Private Sub txtParty_GotFocus()
   If PopUpValue1 <> "" Then
      txtParty.text = PopUpValue1
   End If
End Sub
Private Sub txtParty_LostFocus()
  PopUpValue1 = ""
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
   If Val(txtQty.text) = 0 Then
      txtQty.SetFocus
      Exit Sub
   End If
   If cmdSave.Enabled = True Then
      cmdSave.SetFocus
   End If
   End If
End Sub
Private Sub txtRecno_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    RecDates.SetFocus
 End If
End Sub
Private Sub txtRem_LostFocus()
  
If RS.State = 1 Then RS.close
RS.Open "select * from pass where pass='" & cp & "'", con
If RS.EOF = True Then
MsgBox "Enter Valid Password !!", vbInformation
Exit Sub
End If
  
If cboParty.text <> "" Then
If MsgBox("Want To Change Remarks ?", vbQuestion + vbYesNo) = vbYes Then
   'con.Execute "update ReceiveIssueParty set Remarks = '" & txtRem.Text & "' where PartyName='" & cboParty.Text & "'"
   con.Execute "update sledger set PartyRemarks = '" & txtrem.text & "' where subledger='" & cboParty.text & "'"

End If
End If

End Sub
Private Sub Unautho_Click()
If Unautho.value = True Then
    Call cmdshow_Click
End If
End Sub
Public Sub updateDonnation(scid_ As String, donid_ As String, rowid As String)


If ch_ = False Then

    Dim fdateSale, fdateSaleRet
    If RS.State = 1 Then RS.close
    RS.Open "select fromDate,toDate,fromDateSRet,toDateSRet,NotCreated,Current_Next from turnOverDis order by Current_Next", CCON
    If RS.EOF = False Then
       fdateSale = RS!fromdate
       fdateSaleRet = RS!fromDateSRet
       dt_strR = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromDateSRet & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDateSRet & "',103))"
       
       RS.MoveNext
       If RS!current_next = "Next" And RS!NotCreated <> "y" Then
          con.Execute "exec spnetsale '" & RS!fromdate & "','" & RS!todate & "','" & RS!fromDateSRet & "','" & RS!toDateSRet & "'"
       End If
       
    End If
    
    If RS.State = 1 Then RS.close
    RS.Open "select fromDate,toDate,fromDateSRet,toDateSRet from turnOverDis where (Current_Next='next' and NotCreated='y')", CCON
    If RS.EOF = False Then
       dt_strR = "(INVOICEDATE >= convert(smalldatetime,'" & fdateSaleRet & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDateSRet & "',103))"
       con.Execute "exec spnetsale '" & fdateSale & "','" & RS!todate & "','" & fdateSaleRet & "','" & RS!toDateSRet & "'"
    End If

    ch_ = True

End If


'======================================
Dim qty_sale, qty_don
Dim ss_ As New ADODB.Recordset
Dim scid, gross_net As String
Dim Sp_per, Adj_per, finalAmt, AdvAmt As Double

'---------------------------------------------------------------------------
scid = scid_
qty_don = 0
finalAmt = 0
Sp_per = 0
Adj_per = 0
If ss_.State = 1 Then ss_.close
ss_.Open "SELECT finalAmt,SponsorshipOn,Sponsorship_per,ReturnAdj,AdvAmt  FROM DonnationMain where (ScID='" & scid & "') group by finalAmt,SponsorshipOn,Sponsorship_per,ReturnAdj,AdvAmt", con
If ss_.RecordCount > 0 Then
   If ss_!SponsorshipOn = "Gross" Then
      gross_net = "gross"
   Else
      gross_net = "net"
   End If
   Sp_per = ss_!Sponsorship_per
   Adj_per = ss_!ReturnAdj
   finalAmt = ss_!finalAmt
   AdvAmt = ss_!AdvAmt
End If

'If rs1.State = 1 Then rs1.close
'rs1.Open "select sum(AdvAmt) from DonnationMain where (ScID='" & scid & "')", con
'If Not IsNull(rs1(0)) Then
'   AdvAmt = rs1(0)
'Else
'   AdvAmt = 0
'End If

'---------------------------------------------------------------------------

qty_sale = 0
scid = scid_

If ss_.State = 1 Then ss_.close
ss_.Open "SELECT sum(AMOUNT),sum(NETAMOUNT) FROM tmpinvoiceBQry where ScID='" & scid & "'", con
If Not IsNull(ss_(0)) Then
   If gross_net = "gross" Then
      qty_sale = ss_(0)
   Else
      qty_sale = ss_(1)
   End If
End If
'================================================
afterAdj = qty_sale
If Adj_per > 0 Then
   afterAdj = Round(qty_sale - Round((qty_sale * Adj_per / 100), 0), 0)
Else
   afterAdj = Round(qty_sale, 0)
End If

If Sp_per > 0 Then
   afterAdj = Round((afterAdj * Sp_per / 100), 0)
End If
'================================================
If AdvAmt > 0 Then
   afterAdj = afterAdj - AdvAmt
End If


qty_sale = afterAdj


'If qty_sale < 0 Then
   con.Execute "update DonnationMain set tobeupdate=" & qty_sale & " where dno=" & donid_ & ""
   DoEvents
   vs.TextMatrix(rowid, 8) = "Round Of Amount : " & qty_sale
   DoEvents
   DoEvents
   If qty_sale < 0 Then
   For k1 = 0 To 8
    vs.Cell(flexcpBackColor, rowid, k1) = vbGreen
    DoEvents
   Next
   End If
'Else
'   con.Execute "update DonnationMain set tobeupdate='' where dno=" & donid_ & ""
'End If

End Sub
Private Sub vs_DblClick()

On Error GoTo aa_

If vs.Col = 8 Then

Dim scid_ As String
Dim donid_ As String

Screen.MousePointer = vbHourglass
donid_ = vs.TextMatrix(vs.RowSel, 1)
scid_ = Mid(vs.TextMatrix(vs.RowSel, 3), InStr(vs.TextMatrix(vs.RowSel, 3), ":") + 1)
updateDonnation scid_, donid_, vs.RowSel
Screen.MousePointer = Default

End If

Exit Sub
aa_:
Screen.MousePointer = Default
MsgBox "" & err.DESCRIPTION

End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
Screen.MousePointer = vbHourglass
If KeyCode = 13 Then
If sales.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         invoice.Show  '    sales
   End If
ElseIf cash.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         countersale.Show  '
   End If
ElseIf credit.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
        ' MainMenu.Toolbar1.Visible = False
         Critnote.Show
   End If
ElseIf crdit.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         Creditnotefile.Show
   End If
ElseIf dbit.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         ''MainMenu.Toolbar1.Visible = false
         Debitnotefile.Show
   End If
End If
End If
Screen.MousePointer = vbDefault
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If vs.Col = 4 Then
   sendkeys "{down}"
End If
End If
End Sub
Sub searchSundry_Creditor()

If RS.State = 1 Then RS.close
RS.Open "select YEAROPENING from SLEDGER where " & stringyear & " and SUBLEDGER='" & cboParty.text & "'", con
If RS.EOF = False Then
txtOp.text = Format(RS.Fields(0).value, "0.00")

If RS.Fields(0).value > 0 Then
   cboop.text = "Dr"
Else
   cboop.text = "Cr"
End If
End If

txtOp.text = Abs(txtOp.text)

'========================================================
Dim ff As New ADODB.Recordset
If ff.State = 1 Then ff.close
ff.Open "select VoucherType,VoucherNumber,VoucherDate,DESCRIPTION,Amount,DebitorCredit from VOUCHERS where " & stringyear & " and SubLedger='" & cboParty & "' order by VoucherDate,VoucherNumber", con
vs1.rows = ff.RecordCount + 1
For J = 1 To vs1.rows - 1
 If ff.EOF = False Then
     vs1.TextMatrix(J, 0) = ff.Fields(0).value
     vs1.TextMatrix(J, 1) = ff.Fields(1).value
     vs1.TextMatrix(J, 2) = ff.Fields(2).value
     vs1.TextMatrix(J, 3) = ff.Fields(3).value & ""
     
     If ff!DebitorCredit = "D" Then
       vs1.TextMatrix(J, 4) = Format(ff.Fields(4).value, "0.00")
       vs1.TextMatrix(J, 5) = Format(0, "0.00")
     Else
       vs1.TextMatrix(J, 5) = Format(ff.Fields(4).value, "0.00")
       vs1.TextMatrix(J, 4) = Format(0, "0.00")
     End If
     
      ff.MoveNext
 End If
Next



'===================
    If RS.State = 1 Then RS.close
    RS.Open "select cnn,cnd,na,todid,toddate,CNCategory from Cnf1a where  " & stringyear & " and psld='" & cboParty.text & "'", con
    If RS.EOF = False Then
     vs1.rows = vs1.rows + RS.RecordCount
     For J = J To vs1.rows - 1
      
     s10 = ""
      
     If rs1.State = 1 Then rs1.close
     rs1.Open "select narr from CreditNotDet where cnn='" & RS.Fields("cnn").value & "'", con
     While rs1.EOF = False
        If s10 = "" Then
           s10 = rs1!NARR
        Else
           s10 = s10 & ", " & rs1!NARR
        End If
     rs1.MoveNext
     Wend
      
     If RS.EOF = False Then
    
      vs1.TextMatrix(J, 0) = "CN"
      vs1.TextMatrix(J, 1) = RS.Fields("cnn").value
      vs1.TextMatrix(J, 2) = RS.Fields("cnd").value
      vs1.TextMatrix(J, 3) = "Credit Note -" & s10
      vs1.TextMatrix(J, 5) = Format(RS.Fields("na").value, "0.00")
      vs1.TextMatrix(J, 4) = 0
      
      If Not IsNull(RS!todid) Then
         If RS!CNCategory = "Adjustment" Then
            vs1.TextMatrix(J, 10) = RS.Fields("todid").value & "-" & RS.Fields("toddate").value
         End If
      End If

      
      RS.MoveNext
    
    End If
    
    Next
    End If
     
     
     
    '===================
    If RS.State = 1 Then RS.close
    RS.Open "select dnn,dnd,psld,na,n from dnfa where  " & stringyear & " and psld='" & cboParty.text & "'", con
    If RS.EOF = False Then
     vs1.rows = vs1.rows + RS.RecordCount
     For J = J To vs1.rows - 1
     
     s10 = ""
    
     If RS.EOF = False Then
    
      If (Not IsNull(RS!n) Or RS!n <> "") Then s10 = "" & RS!n
     
     '----------------------------------
       If rs1.State = 1 Then rs1.close
       rs1.Open "select narr from debitNotDet where dnn='" & RS.Fields("dnn").value & "'", con
       If rs1.EOF = False Then s10 = ""
       While rs1.EOF = False
         If s10 = "" Then
           s10 = rs1!NARR
         Else
           s10 = s10 & ", " & rs1!NARR
         End If
        rs1.MoveNext
       Wend
     '------------------------------------
      
        
      vs1.TextMatrix(J, 0) = "DN"
      vs1.TextMatrix(J, 1) = RS.Fields("dnn").value
      vs1.TextMatrix(J, 2) = RS.Fields("dnd").value
      vs1.TextMatrix(J, 3) = "Debit Note " & s10
      vs1.TextMatrix(J, 4) = Format(RS.Fields("na").value, "0.00")
      vs1.TextMatrix(J, 5) = 0
      RS.MoveNext
    End If
    Next
    End If
 


CalculateTotalDrCr


vs1.FormatString = "^VType|^Bill|^Dates|Description|>Dr|>Cr|Balance|Dr/Cr"
vs1.ColWidth(0) = 600
vs1.ColWidth(1) = 700
vs1.ColWidth(2) = 1000
vs1.ColWidth(3) = 6200
vs1.ColWidth(4) = 1350
vs1.ColWidth(5) = 1350
vs1.ColWidth(6) = 1200
vs1.ColWidth(7) = 600
vs1.ColWidth(8) = 0
vs1.ColWidth(9) = 0
vs1.ColWidth(10) = 0




For h1 = 1 To 8
vs1.Cell(flexcpFontSize, 0, h1) = 11
Next
    
   
   
   
   
   
   vs1.WordWrap = True


End Sub
Sub CalculateTotalDrCr()
    
'On Error Resume Next

On Error GoTo aa1

Dim Balance As Long
Dim dr1, cr1, prbal
Dim rs_1 As New ADODB.Recordset

Dim Str
Str = ""
dr1 = 0
cr1 = 0
txtClosing.text = 0
txtcr.text = 0
If RS.State = 1 Then RS.close
RS.Open "select Op,drcr,YEAROPENING from SLEDGER where " & stringyear & " and SUBLEDGER='" & cboParty.text & "'", con
If RS.EOF = False Then

If lblCr = "cr" Then
   txtOp.text = Format(RS.Fields(2).value, "0.00")
   cmdSave.Enabled = False
Else
   txtOp.text = Format(RS.Fields(0).value, "0.00")
   cmdSave.Enabled = True
End If

txtOp.text = Abs(txtOp.text)

If lblCr = "dr" Then

    If Len(RS.Fields("drcr").value) >= 2 Then
        If UCase(RS.Fields("drcr").value) = UCase("dr") Then
          cboop.text = "Dr"
        Else
          cboop.text = "Cr"
        End If
    End If
    
    Else
       'txtOp.Text = 0
    End If

End If






'=====================================================
  If vs1.rows <= 1 Then
        txtBalance = txtOp
        closingcr = cboop
     Exit Sub
  End If
'=====================================================


If cboop.text = "Dr" Then
   dr1 = (Val(txtOp.text) + Val(vs1.TextMatrix(1, 4)))
   cr1 = Val(vs1.TextMatrix(1, 5))
Else
   cr1 = (Val(txtOp.text) + Val(vs1.TextMatrix(1, 5)))
   dr1 = Val(vs1.TextMatrix(1, 4))
End If


prbal = dr1 - cr1
If prbal < 0 Then
vs1.TextMatrix(1, 6) = Format(-1 * prbal, "0.00")
vs1.TextMatrix(1, 7) = "Cr"
Else
vs1.TextMatrix(1, 6) = Format(prbal, "0.00")
vs1.TextMatrix(1, 7) = "Dr"
End If


For I = 1 To vs1.rows - 1
If vs1.TextMatrix(I, 0) <> "" Then
txtClosing.text = (Val(txtClosing.text) + Val(vs1.TextMatrix(I, 4)))
txtcr.text = (Val(txtcr.text) + Val(vs1.TextMatrix(I, 5)))
'-----Balance---------------
If I >= 2 Then
dr1 = Val(vs1.TextMatrix(I, 4))
cr1 = (-1 * Val(vs1.TextMatrix(I, 5)))
bal = dr1 + cr1
If Str = "Cr" Then
bal = prbal + bal
Else
bal = prbal + bal
End If
If bal < 0 Then
vs1.TextMatrix(I, 6) = Format(-1 * bal, "0.00")
vs1.TextMatrix(I, 7) = "Cr"
Else
vs1.TextMatrix(I, 6) = Format(bal, "0.00")
vs1.TextMatrix(I, 7) = "Dr"
End If
prbal = bal
Str = vs1.TextMatrix(I, 7)
End If


If (cboParty.text = "O2020 ONLINE SALES (PAY U)" Or cboParty.text = "O2020 ONLINE SALES (PAY U)" Or cboParty.text = "A2020 AMAZON.IN") Then
   GoTo aaa:
End If

'--------Donation Details-----------------------------------------------------------
        
        d11 = ""
        str11 = "SELECT distinct DonnationMain.DNo,DonnationMain.DDate,DonnationMainDet.INVOICENO,DonnationMainDet.Godown FROM  DonnationMain INNER JOIN" & _
              " DonnationMainDet ON DonnationMain.DNo = DonnationMainDet.DNo where DonnationMainDet.fyear='" & session & "' and DonnationMainDet.INVOICENO=" & vs1.TextMatrix(I, 1) & " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & SessionLastDate1 & "' ,103)"
        ss1_ = ""
        If rs_1.State = 1 Then rs_1.close
        rs_1.Open str11, con
        While rs_1.EOF = False
            If rs_1!Godown = "I" Then
               d11 = "I"
            Else
               d11 = "CI"
            End If
            If vs1.TextMatrix(I, 0) = d11 Then
                If ss1_ = "" Then
                   ss1_ = rs_1(0) & "-" & rs_1(1)
                Else
                   ss1_ = ss1_ & "," & rs_1(0) & "-" & rs_1(1)
                End If
            End If
            rs_1.MoveNext
        Wend
        If ss1_ <> "" Then
           vs1.TextMatrix(I, 8) = ss1_
        End If

'-----End Donation Details-------------------------------------------------------------
    
            d11 = ""
            str11 = " SELECT distinct SalesAdjustment.DNo,SalesAdjustment.DDate,SalesAdjustmentDet.Godown FROM SalesAdjustment INNER JOIN" & _
                  " SalesAdjustmentDet ON SalesAdjustment.DNo = SalesAdjustmentDet.DNo where SalesAdjustmentDet.fyear='" & session & "' and SalesAdjustmentDet.INVOICENO=" & vs1.TextMatrix(I, 1) & " and SalesAdjustment.SCId='" & cboParty & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & SessionLastDate1 & "' ,103)"
            ss1_ = ""
            If rs_1.State = 1 Then rs_1.close
            rs_1.Open str11, con
            
            While rs_1.EOF = False
                If rs_1!Godown = "I" Then
                   d11 = "I"
                Else
                   d11 = "CI"
                End If
                If vs1.TextMatrix(I, 0) = d11 Then
                    If ss1_ = "" Then
                       ss1_ = rs_1(0) & "-" & rs_1(1)
                    Else
                       ss1_ = ss1_ & "," & rs_1(0) & "-" & rs_1(1)
                    End If
                End If
                rs_1.MoveNext
            Wend
            If ss1_ <> "" Then
               vs1.TextMatrix(I, 9) = ss1_
            End If
            
End If

'-----End Adjustment Details--------------------------------------------------------

aaa:


Next

'====================================================================================
'For Last Data where Donnation & sale Adj is create ----------------------------------
'====================================================================================
If rs_don.State = 1 Then rs_don.close
rs_don.Open "select LastDatabase from data", CCON
If rs_don.EOF = False Then
    Set con_don = New ADODB.Connection
    If LCase(server_) = "server" Then
        con_don.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs_don!LastDatabase & "; UID=" & sql_user & "; PWD=" & sql_pass
        
        con_don.Open
    Else
       con_don.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=.; DATABASE=" & rs_don!LastDatabase & "; UID=; PWD=;"
       con_don.Open
    End If
End If



If (cboParty.text = "O2020 ONLINE SALES (PAY U)" Or cboParty.text = "O2020 ONLINE SALES (PAY U)" Or cboParty.text = "A2020 AMAZON.IN") Then
   GoTo aaa1:
End If


For I = 1 To vs1.rows - 1
    '--------Donation Details-----------------------------------------------------------
    
    If (vs1.TextMatrix(I, 1) <> "") Then
    
    d11 = ""
    str11 = "SELECT distinct DonnationMain.DNo,DonnationMain.DDate,DonnationMainDet.INVOICENO,DonnationMainDet.Godown FROM  DonnationMain INNER JOIN" & _
          " DonnationMainDet ON DonnationMain.DNo = DonnationMainDet.DNo where DonnationMainDet.fyear='" & session & "' and DonnationMainDet.INVOICENO=" & vs1.TextMatrix(I, 1) & " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & SessionLastDate1 & "' ,103)"
    ss1_ = ""
    If rs_1.State = 1 Then rs_1.close
    rs_1.Open str11, con_don
    While rs_1.EOF = False
        If rs_1!Godown = "I" Then
           d11 = "I"
        Else
           d11 = "CI"
        End If
        If vs1.TextMatrix(I, 0) = d11 Then
            If ss1_ = "" Then
               ss1_ = rs_1(0) & "-" & rs_1(1)
            Else
               ss1_ = ss1_ & "," & rs_1(0) & "-" & rs_1(1)
            End If
        End If
        rs_1.MoveNext
    Wend
    If ss1_ <> "" Then
       vs1.TextMatrix(I, 8) = ss1_
    End If
    
    End If
    
    '-----End Donation Details-------------------------------------------------------------
        
    If (vs1.TextMatrix(I, 1) <> "") Then
    
    d11 = ""
    str11 = " SELECT distinct SalesAdjustment.DNo,SalesAdjustment.DDate,SalesAdjustmentDet.Godown FROM SalesAdjustment INNER JOIN" & _
          " SalesAdjustmentDet ON SalesAdjustment.DNo = SalesAdjustmentDet.DNo where SalesAdjustmentDet.fyear='" & session & "' and SalesAdjustment.SCId='" & cboParty & "' and SalesAdjustmentDet.INVOICENO=" & vs1.TextMatrix(I, 1) & " and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & SessionLastDate1 & "' ,103)"
    ss1_ = ""
    If rs_1.State = 1 Then rs_1.close
    rs_1.Open str11, con_don
    While rs_1.EOF = False
        If rs_1!Godown = "I" Then
           d11 = "I"
        Else
           d11 = "CI"
        End If
        If vs1.TextMatrix(I, 0) = d11 Then
            If ss1_ = "" Then
               ss1_ = rs_1(0) & "-" & rs_1(1)
            Else
               ss1_ = ss1_ & "," & rs_1(0) & "-" & rs_1(1)
            End If
        End If
        rs_1.MoveNext
    Wend
    If ss1_ <> "" Then
       vs1.TextMatrix(I, 9) = ss1_
    End If
            
    End If
'-----End Adjustment Details--------------------------------------------------------
Next


aaa1:
'================================================================================================

txtClosing.text = Format(txtClosing.text, "0.00")
sum11 = 0

txtcr.text = Format(txtcr.text, "0.00")
If cboop.text = "Dr" Then
  txtClosing.text = Format((CDbl(txtClosing.text)), "0.00")
Else
  txtcr.text = Format((CDbl(txtcr.text)), "0.00")
End If

txtBalance.text = prbal
closingcr.text = Str



txtBalance.text = Format(txtBalance.text, "0.00")
lblTotalRecord.Caption = "Tot.Rows : " & vs1.rows


Exit Sub
aa1:
MsgBox "" & err.DESCRIPTION, vbCritical


End Sub
Sub SaveDatainTempledger()

Dim V

con.Execute "delete  from templedger1 where " & stringyear & ""
For I = 1 To vs1.rows - 1
If vs1.TextMatrix(I, 1) <> "" Then
    
   con.Execute "INSERT INTO  templedger1(dates,Billtype,Bill,Des,Dr,Cr,Balance,drcr,fyear,setupid,repname)  values('" & Format(vs1.TextMatrix(I, 2), "MM/dd/yyyy") & "','" & vs1.TextMatrix(I, 0) & "', " & vs1.TextMatrix(I, 1) & ",'" & vs1.TextMatrix(I, 3) & "' ," & vs1.TextMatrix(I, 4) & "," & vs1.TextMatrix(I, 5) & "," & Val(vs1.TextMatrix(I, 6)) & ",'" & vs1.TextMatrix(I, 7) & "','" & session & "'," & setupid & ",'" & vs1.TextMatrix(I, 10) & "')"
   
End If
Next

Dim ff As New ADODB.Recordset
If ff.State = 1 Then ff.close
ff.Open "select Billtype,bill,dates,des,dr,cr,Balance,drcr,repname from templedger1  where " & stringyear & " order by dates,bill", con
vs1.rows = ff.RecordCount + 1
For J = 1 To vs1.rows - 1
 If ff.EOF = False Then
     vs1.TextMatrix(J, 0) = ff.Fields(0).value
     vs1.TextMatrix(J, 1) = ff.Fields(1).value
     vs1.TextMatrix(J, 2) = ff.Fields(2).value
     vs1.TextMatrix(J, 3) = ff.Fields(3).value
     vs1.TextMatrix(J, 4) = Format(ff.Fields(4).value, "0.00")
     vs1.TextMatrix(J, 5) = Format(ff.Fields(5).value, "0.00")
     vs1.TextMatrix(J, 6) = Format(ff.Fields(6).value, "0.00")
     vs1.TextMatrix(J, 10) = ff.Fields("repname").value & ""
     
     If vs1.TextMatrix(J, 0) = "I" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select app_add,appno from invoicea where (invoiceno=" & ff.Fields(1).value & " and app_add='y' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" & fdate1_ & "',103))", con
        If rs1.EOF = False Then
           vs1.TextMatrix(J, 11) = rs1.Fields("appno").value & ""
           vs1.Cell(flexcpBackColor, J, 0) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 1) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 2) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 3) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 4) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 5) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 6) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 7) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 8) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 9) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 10) = &H78CFE9
           vs1.Cell(flexcpBackColor, J, 11) = &H78CFE9
        End If
     End If
     
     ff.MoveNext
 End If
Next
End Sub
Sub seriesWiseDiscount()
  
    Dim rs_dis As New ADODB.Recordset
    Dim lastYrs As String
    
    Set rs_dis = New ADODB.Recordset
    
    lastYrs = ""
    
    If (DateDiff("d", Now, SessionLastDate) <= 0) Then
       lastYrs = "current"
    Else
       lastYrs = "last"
    End If
  
  
  
  
  If lastYrs = "last" Then
     rs_dis.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(cboParty.text, 1, 5) & "'", con_LAST
  Else
     rs_dis.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(cboParty.text, 1, 5) & "'", con
  End If
  
  
   
  
  
  If rs_dis.EOF = False Then
        Label1(1).Visible = True
        Label1(2).Visible = True
        Timer2.Enabled = True
  Else
        Label1(1).Visible = False
        Label1(2).Visible = False
        Timer2.Enabled = False
  End If

End Sub
Private Sub cboParty_GotFocus()

Dim ph_rs As New ADODB.Recordset
cboParty.BackColor = &HFFFFC0

If search_v = False Then

I = 1
If PopUpValue3 = "" Then
   PopUpValue2 = cboParty.text
    
    If ph_rs.State = 1 Then ph_rs.close
    ph_rs.Open "select top 1 profile_ from sledger where " & stringyear & " and subledger='" & cboParty.text & "'", con
    If ph_rs.EOF = False Then

        If UCase(ph_rs(0)) = UCase("CASH PARTY") Then
          Label1(20).Visible = True
          Label1(21).Visible = True
          
          Label1(20).Caption = "CASH PARTY"
          Label1(21).Caption = "CASH PARTY"
        
          Timer1.Enabled = True
        Else
          Label1(20).Visible = False
          Label1(21).Visible = False
          Timer1.Enabled = False
        End If
        
    End If
   
End If

If PopUpValue3 <> "" Then
cboParty.text = PopUpValue3

'' change

If LCase(server_) = "server" Then
fillDocument (Mid(cboParty.text, 1, 5))
End If

seriesWiseDiscount
lblfrt.Caption = ""

Set ph_rs = New ADODB.Recordset
ph_rs.Open "select phone,PartyRemarks,MOBILE,profile_,freight from sledger where " & stringyear & " and subledger='" & cboParty.text & "'", con




If ph_rs.EOF = False Then
   
        phone.Caption = ph_rs(0) & "," & ph_rs!mobile
        
        lblfrt.Caption = "Freight : " & ph_rs!freight & ""
        
        txtrem.text = ph_rs.Fields("PartyRemarks").value & ""
        
        CASH_profile = ph_rs!profile_ & ""
        
   
      If UCase(ph_rs!profile_) = UCase("CASH PARTY") Then
        Label1(20).Visible = True
        Label1(21).Visible = True
        
        Label1(20).Caption = "CASH PARTY"
        Label1(21).Caption = "CASH PARTY"

        Timer1.Enabled = True
      Else
        Label1(20).Visible = False
        Label1(21).Visible = False
        Timer1.Enabled = False
      End If
      
      If UCase(ph_rs!profile_) = UCase("NO DEALING") Then
        Label1(20).Visible = True
        Label1(21).Visible = True
        Label1(20).Caption = "NO DEALING"
        Label1(21).Caption = "NO DEALING"
        Timer1.Enabled = True
      
      End If
   
Else
   phone.Caption = ""
   txtrem.text = ""
End If
End If

'TOD----------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------------------

I = 1

If PopUpValue2 = "" Then
   PopUpValue2 = cboParty.text
End If


If PopUpValue2 <> "" Then


cboParty.text = PopUpValue2
Set ph_rs = New ADODB.Recordset
ph_rs.Open "select phone,PartyRemarks,MOBILE from sledger where " & stringyear & " and subledger='" & cboParty.text & "'", con
If ph_rs.EOF = False Then
   
   phone.Caption = ph_rs(0) & "," & ph_rs!mobile
   txtrem.text = ph_rs.Fields("PartyRemarks").value & ""
   
Else
   
   phone.Caption = ""
   txtrem.text = ""
  
End If




End If
End If



End Sub
Private Sub cboParty_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dr, CR As Double




If KeyCode = 114 Then
  search_v = True
Else
  search_v = False
End If


If search_v = False Then

If KeyCode = 113 Then
    
    searchType = "party"
    
    lblCr = "dr"
    value = "select distinct(Party),Code,subledger from SLEDGER where (gledger='SUNDRY DEBTORS') and " & stringyear & "  order by party"
    popuplist_client value, CCON
    set_focus = True
    
    
    
End If

If KeyCode = 115 Then
    
    searchType = "ledger"
    lblCr = "cr"
    value = "select distinct(Party),Code,subledger from SLEDGER where (gledger='SUNDRY CREDITORS' or gledger='IMPREST A/C') and " & stringyear & "  order by party"
    popuplist_client value, CCON
    
    set_focus = True
    
End If




If KeyCode = 13 Then



If cboParty.text = "" Then
  cboParty.SetFocus
  Exit Sub
End If



If lblCr = "cr" Then
   searchSundry_Creditor
Else
   
   If Check3_newledger.value = 1 Then
      Command5_Click
      Exit Sub
   End If
   
   
   dataSearchingrid
End If





cmdPrint.Enabled = True

dr = 0
CR = 0

For I = 1 To vs1.rows - 1
  dr = dr + Val(vs1.TextMatrix(I, 4))
  CR = CR + Val(vs1.TextMatrix(I, 5))
Next

drLebel.Caption = Format(dr, "0.00")
CrLebel.Caption = Format(CR, "0.00")
    
txtdes.SetFocus

    
End If

'========================================================================================
Else
'========================================================================================


If KeyCode = 114 Then
   
     
   
   value = "SELECT  trim(mid( SLEDGER.SUBLEDGER,instr(SLEDGER.SUBLEDGER,',')+1)) as city,SUBLEDGER," & _
    "Code FROM SLEDGER where instr(SLEDGER.SUBLEDGER,',')>0 and gledger='SUNDRY DEBTORS' and " & stringyear & " order by trim(mid( SLEDGER.SUBLEDGER,instr(SLEDGER.SUBLEDGER,',')+1))"
    
    popuplist_client value, CCON
    
    set_focus = True
End If

If KeyCode = 13 Then
If cboParty.text = "" Then
  cboParty.SetFocus
  Exit Sub
End If




dataSearchingrid
cmdPrint.Enabled = True

dr = 0
CR = 0

For I = 1 To vs1.rows - 1
  dr = dr + Val(vs1.TextMatrix(I, 4))
  CR = CR + Val(vs1.TextMatrix(I, 5))
Next

drLebel.Caption = Format(dr, "0.00")
CrLebel.Caption = Format(CR, "0.00")
    
txtdes.SetFocus

End If

End If



End Sub
Sub dataSearchingrid()
Screen.MousePointer = vbHourglass
I = 1



If PopUpValue3 <> "" Then
   vs1.Clear
   vs1.rows = 1
   fillGrid
End If



If cboParty.text <> "" Then
   
   If RS.State = 1 Then RS.close
   RS.Open "select YEAROPENING from SLEDGER where " & stringyear & " and SUBLEDGER='" & cboParty.text & "'", con
   If RS.EOF = False Then
      txtOp.text = Format(RS.Fields(0).value, "0.00")
      If RS.Fields(0).value > 0 Then
         cboop.text = "Dr"
      Else
         cboop.text = "Cr"
      End If
   End If

   
   
   SaveDatainTempledger
   CalculateTotalDrCr
End If



setWidth
PopUpValue1 = ""
Screen.MousePointer = vbDefault


End Sub
Private Sub cboParty_LostFocus()





cboParty.BackColor = &HFFFFFF
PopUpValue1 = ""
PopUpValue3 = ""
PopUpValue2 = ""



End Sub
Sub DelFunction()
    Dim Del As New ADODB.Recordset
    If Del.State = 1 Then Del.close
    Set Del = con.Execute("delete from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.text & " and firm='" & firm & "'")
End Sub
Private Sub cmdDel_Click()
  If RS.State = 1 Then RS.close
  RS.Open "select * from pass where pass='" & cp & "'", con
  If RS.EOF = True Then
     MsgBox "Enter Valid Password !!", vbInformation
     Exit Sub
  End If
  
   If MsgBox("Do U Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
       DelFunction
       fillGrid
       dataSearchingrid
       Call cmdRefresh_Click
       cmdModify.Enabled = False
       cmdDel.Enabled = False
   End If
End Sub
Private Sub cmdMain_Click()
If strledger = "cp" Then
If Val(txtQty.text) > 0 And txtdes.text <> "" And cboParty.text <> "" Then
   If MsgBox("Want To Save & Exit ?", vbQuestion + vbYesNo) = vbYes Then
          SaveMain
          Call cmdRefresh_Click
          fillGrid
          cmdModify.Enabled = False
          cmdDel.Enabled = False
          cboParty.SetFocus
          dataSearchingrid
          Unload Me
          Exit Sub
   End If
End If
End If
If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
  Unload Me
End If
End Sub
Sub setWidth()
    
   Dim d1  As Integer
    Dim rss_ As New ADODB.Recordset
    
    d1 = 0
    
      
      If donnation_visible = "n" Then
        vs1.FormatString = "^VType|^Bill|^Dates|Description|>Dr|>Cr|Balance|Dr/Cr||Adjustm./CrNO|TOD Details|App.NO"
        d1 = 0
     Else
        vs1.FormatString = "^VType|^Bill|^Dates|Description|>Dr|>Cr|Balance|Dr/Cr|Extra Dis.|Adjustm./CrNo|TOD Details|App.NO"
        d1 = 1100
     End If
     
     
    If RS.State = 1 Then RS.close
    RS.Open "select top 1 dno from DonnationMain", con
    If RS.RecordCount = 0 Then
        vs1.FormatString = "^VType|^Bill|^Dates|Description|>Dr|>Cr|Balance|Dr/Cr||Adjustm./CrNO|TOD Details|App.NO"
        d1 = 0
    End If
     
      
    vs1.ColWidth(0) = 450
    vs1.ColWidth(1) = 700
    vs1.ColWidth(2) = 1000
    vs1.ColWidth(3) = 4000
    vs1.ColWidth(4) = 1100
    vs1.ColWidth(5) = 1250
    vs1.ColWidth(6) = 1050
    vs1.ColWidth(7) = 500
    vs1.ColWidth(8) = d1
    vs1.ColWidth(9) = 1300
    vs1.ColWidth(10) = 1100
    vs1.ColWidth(11) = 800
    
    
    For h1 = 1 To 9
    vs1.Cell(flexcpFontSize, 0, h1) = 11
    Next
   
   
   
   vs1.WordWrap = True
    
    
   DoEvents

End Sub
Private Sub cmdModify_Click()
Set RS = New ADODB.Recordset
RS.Open "select * from pass where pass='" & cp & "'", con
If RS.EOF = True Then
MsgBox "Enter Valid Password !!", vbInformation
Exit Sub
End If


On Error GoTo aa1

If MsgBox("Do U Want To Update ?", vbQuestion + vbYesNo) = vbYes Then
'DelFunction
con.Execute "update ReceiveIssueParty set Dr=0,cr=0 where RecNo=" & txtRecno.text & " and firm='" & firm & "'"

'------------------------
Set RS = New ADODB.Recordset
RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.text & " and firm='" & firm & "'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
    RS.Fields("RecNo").value = txtRecno.text
    RS.Fields("Dates").value = RecDates.value
    RS.Fields("PartyName").value = cboParty.text
    RS.Fields("Particullar").value = txtdes.text
    If Receive.value = True Then
        RS.Fields("Dr").value = Val(txtQty.text)
    Else
        RS.Fields("Cr").value = Val(txtQty.text)
    End If
    RS.Fields("Remarks").value = Trim(txtrem.text)
RS.update
End If
'------------------------



fillGrid
CalculateTotalDrCr
setWidth
Call cmdRefresh_Click
vs1.SetFocus

For I = 1 To vs1.rows - 1
sendkeys "{down}"
Next

cmdModify.Enabled = False
cmdDel.Enabled = False
End If

Exit Sub
aa1:
MsgBox "" & err.DESCRIPTION, vbCritical

End Sub
Private Sub cmdRefresh_Click()
 

 
 Dim o As Object
 
 CASH_profile = ""
 txtchequeeNo.text = ""
 txtQty.text = ""
 lblTOD.Caption = ""
 frmOrderList.Visible = False
 
 set_focus = False
 List1_ch.Visible = False
 
 con.Execute "delete from templedger1"
 
 maxId
 cmdModify.Enabled = False
 cmdDel.Enabled = False
 RecDates.SetFocus
 
 If lblCr = "cr" Then
 Else
 cmdSave.Enabled = True
 End If
 
Label1(20).Visible = False
Label1(21).Visible = False
Timer1.Enabled = False
 
'con.Execute "delete from tmpDonnationnew where uid='" & Sys_user_ & "'"
'con.Execute "delete from tmpSaladjust where uid='" & Sys_user_ & "'"
'con.Execute "delete from templedger6 where userid='" & Sys_user_ & "'"
 
 
 Screen.MousePointer = vbDefault
 bb2 = False

End Sub
Private Sub cmdSave_Click()

On Error GoTo aa:

If cboParty.text = "" Then
MsgBox "Please Select Party Name !!", vbInformation
Exit Sub
End If

If (txtQty.text = "" Or txtQty.text = "0") Then
   MsgBox "Please Enter Amount!!", vbInformation
   txtQty.SetFocus
   Exit Sub
End If

If Val(txtQty.text) = 0 Then
   MsgBox "Please Enter Amount!!", vbInformation
   txtQty.SetFocus
   Exit Sub
End If


If RS.State = 1 Then RS.close
RS.Open "select * from pass where pass='" & cp & "'", con
If RS.EOF = True Then
MsgBox "Enter Valid Password !!", vbInformation
Exit Sub
End If

If MsgBox("Do U Want To Save ?", vbInformation + vbYesNo) = vbYes Then

aa1:

SaveMain
cboParty.SetFocus

Call cmdRefresh_Click
fillGrid

cmdModify.Enabled = False
cmdDel.Enabled = False
'----------------
dataSearchingrid
'---------------
 
 
End If
Exit Sub
aa:
maxId
GoTo aa1

End Sub
Sub SaveMain()
   
    maxId
    Set RS = New ADODB.Recordset
    RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.text & " and firm='" & firm & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
       maxId
       RS.AddNew
       RS.Fields("RecNo").value = txtRecno.text
       RS.Fields("Dates").value = RecDates.value
       RS.Fields("PartyName").value = cboParty.text
       RS.Fields("Particullar").value = txtdes.text
       If Receive.value = True Then
          RS.Fields("Dr").value = Val(txtQty.text)
          RS.Fields("Cr").value = 0
        Else
          RS.Fields("Cr").value = Val(txtQty.text)
          RS.Fields("Dr").value = 0
       End If
       
       RS.Fields("firm").value = firm
       RS.Fields("fyear").value = session
       RS.Fields("setupid").value = setupid
       RS.Fields("Remarks").value = txtrem.text
       RS.update
    End If
End Sub
Sub search()

 If lblCr = "cr" Then Exit Sub

 If set_focus = True Then Exit Sub
 
 On Error Resume Next
 
 
 
 If vs1.TextMatrix(vs1.RowSel, 0) = "J" Then
    
    If RS.State = 1 Then RS.close
    RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & vs1.TextMatrix(vs1.RowSel, 1) & " and firm='" & firm & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       txtRecno.text = RS.Fields("RecNo").value
       RecDates.value = RS.Fields("Dates").value
       cboParty.text = RS.Fields("PartyName").value
       txtdes.text = RS.Fields("Particullar").value
       'txtRem.Text = RS.Fields("Remarks").value & ""
       
       If RS.Fields("Dr").value > 0 Then
          Receive.value = True
          txtQty.text = RS.Fields("Dr").value
        Else
          Issue.value = True
          txtQty.text = RS.Fields("Cr").value
       End If
      End If
   cmdSave.Enabled = False
   cmdModify.Enabled = True
   cmdDel.Enabled = True
  ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "C/M" Then
    DoEvents
    vs1.ToolTipText = "dinesh -----------"
    
  ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "CI" Then
    DoEvents
    vs1.ToolTipText = "Chitra Pr"
    
  Else
   cmdModify.Enabled = False
   cmdDel.Enabled = False
   cmdSave.Enabled = True
   txtdes.text = ""
   txtQty.text = ""
  End If
  
End Sub
Private Sub cmdSearch_Click()
Frame1.Visible = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  
  If KeyCode = 27 Then
   If cboPartyList.Visible = True Then
      cboPartyList.Visible = False
      Exit Sub
   End If
  End If
  
  
  
  If KeyCode = 116 Then
  If bb2 = False Then
    vs1.SetFocus
    For I = 1 To vs1.rows - 1
    sendkeys "{down}"
    Next
    bb2 = True
  Else
    Call cmdRefresh_Click
    cboParty.SetFocus
    bb2 = False
  End If
  Exit Sub
  End If
  
  
  
  If KeyCode = 112 Then
     txtdes.SetFocus
     Exit Sub
  End If
   If KeyCode = 27 Then
        If RS.State = 1 Then RS.close
        RS.Open "select * from pass where pass='" & cp & "'", con
        If RS.EOF = True Then
          If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
           Unload Me
           End If
        Exit Sub
        End If
        If Val(txtQty.text) > 0 And txtdes.text <> "" And cboParty.text <> "" Then
        If MsgBox("Want To Save & Exit ?", vbQuestion + vbYesNo) = vbYes Then
            SaveMain
            Call cmdRefresh_Click
            fillGrid
            cmdModify.Enabled = False
            cmdDel.Enabled = False
            cboParty.SetFocus
            dataSearchingrid
            Unload Me
            Exit Sub
        End If
        End If
      If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
         Unload Me
       End If
      ElseIf KeyCode = 13 Then
      ElseIf KeyCode = 113 Then
         kk = False
   End If
End Sub
Sub fillGrid()
    
    
    
    
    Dim fill As New ADODB.Recordset
    Set fill = New ADODB.Recordset
    fill.Open "select RecNo,Dates,Particullar,Dr,Cr,Remarks from ReceiveIssueParty where " & stringyear & " and PartyName='" & cboParty.text & "' and firm='" & firm & "' order by dates,recno", con
    If fill.EOF = False Then
       'txtRem.Text = fill!Remarks & ""
       vs1.rows = fill.RecordCount + 1
       For I = 1 To vs1.rows - 1
           vs1.TextMatrix(I, 0) = "J"
           vs1.TextMatrix(I, 1) = fill.Fields(0).value
           vs1.TextMatrix(I, 2) = fill.Fields(1).value
           vs1.TextMatrix(I, 3) = fill.Fields(2).value
           vs1.TextMatrix(I, 4) = Format(fill.Fields(3).value, "0.00")
           vs1.TextMatrix(I, 5) = Format(fill.Fields(4).value, "0.00")
           vs1.TextMatrix(I, 10) = ""
           fill.MoveNext
       Next
    Else
    vs1.Clear
    End If


    
    '==============
    If firm = "chitra" Then
       SearchFa
    Else
       SearchFa_blueprint
    End If
    '==============
    setWidth
End Sub
Sub maxId()
  Dim rr As New ADODB.Recordset
  Set rr = New ADODB.Recordset
  rr.Open "select max(RecNo) from ReceiveIssueParty where " & stringyear & " and firm='" & firm & "'", con
  If IsNull(rr.Fields(0).value) Then
     txtRecno.text = 1
     Else
     txtRecno.text = rr.Fields(0).value + 1
  End If
End Sub

Private Sub Todate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    from_date = fromdate.value
    to_date = todate.value
    fillGrid
    Frame1.Visible = False
 End If
End Sub

Private Sub txtQty_GotFocus()
   txtQty.BackColor = &HFFFFC0
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
   Dim b As Boolean
   b = val_int(txtQty, KeyAscii)
   If b = False Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtQty_LostFocus()
  txtQty.BackColor = &HFFFFFF
End Sub
Private Sub txtRecno_KeyPress(KeyAscii As Integer)
   On Error Resume Next
  
     bb = val_int(txtRecno, KeyAscii)
     If bb = False Then
        KeyAscii = 0
     End If
  
  If KeyAscii = 13 Then
  
     If RS.State = 1 Then RS.close
     RS.Open "select * from receiveissueparty where " & stringyear & " and recno=" & txtRecno.text & " and firm='" & firm & "'", con
     If RS.EOF = False Then
       
      cmdModify.Enabled = True
      cmdDel.Enabled = True
       
      cboParty.text = RS!partyname
      PopUpValue3 = cboParty.text
      txtrem.text = RS!remarks & ""
      
      RecDates.value = RS.Fields("Dates").value
      txtdes.text = RS.Fields("Particullar").value
       If RS.Fields("Dr").value > 0 Then
          Receive.value = True
          txtQty.text = RS.Fields("Dr").value
      Else
          Issue.value = True
          txtQty.text = RS.Fields("Cr").value
      End If
      dataSearchingrid
     Else
       vs1.Clear
       setWidth
       txtQty.text = ""
       txtdes.text = ""
       cboParty.text = ""
       txtOp.text = ""
       txtBalance.text = ""
     End If
  End If
  
   
  
End Sub
Private Sub txtSlipNo_GotFocus()
 txtSlipNo.BackColor = &HFFFFC0
End Sub
Private Sub txtSlipNo_LostFocus()
txtSlipNo.BackColor = &HFFFFFF
End Sub

Private Sub vs_promotion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If (vs_promotion.Col = 0 Or vs_promotion.Col = 9) Then
     sendkeys "{down}"
  End If
End If
End Sub
Private Sub vs_SelChange()
 If (vs.Col = 7 Or vs.Col = 5) Then
    vs.Editable = flexEDKbdMouse
  Else
    vs.Editable = flexEDNone
 End If
End Sub

Private Sub vs1_Click()
 On Error Resume Next
 
 search
End Sub
Private Sub vs1_DblClick()

On Error Resume Next

set_focus = False

Screen.MousePointer = vbHourglass

Dim s_ As String
Dim n As Integer


If vs1.Col = 8 Then
   
   n = InStr(vs1.TextMatrix(vs1.RowSel, 8), "-")
   inviceNo = Mid(vs1.TextMatrix(vs1.RowSel, 8), 1, n - 1)
   inv_ledger = vs1.TextMatrix(vs1.RowSel, 1)
   pname_ = cboParty.text
   frmDonnation.Show
   
ElseIf vs1.Col = 9 Then
  
  If vs1.TextMatrix(vs1.RowSel, 0) <> "J" Then
   
   n = InStr(vs1.TextMatrix(vs1.RowSel, 9), "-")
   inviceNo = Mid(vs1.TextMatrix(vs1.RowSel, 9), 1, n - 1)
   inv_ledger = vs1.TextMatrix(vs1.RowSel, 1)
   pname_ = cboParty.text
   PopUpValue6 = vs1.TextMatrix(vs1.RowSel, 2)
   frmSalesAdjustment.Show
   
  End If

ElseIf vs1.Col = 10 Then
   n = InStr(vs1.TextMatrix(vs1.RowSel, 10), "-")
   inviceNo = Mid(vs1.TextMatrix(vs1.RowSel, 10), 1, n - 1)
   inv_ledger = vs1.TextMatrix(vs1.RowSel, 1)
   pname_ = cboParty.text
   frmTurnOverDis.Show

ElseIf vs1.Col = 11 Then
   
   'n = InStr(vs1.TextMatrix(vs1.RowSel, 11), "-")
   inviceNo = vs1.TextMatrix(vs1.RowSel, 11)
   inv_ledger = vs1.TextMatrix(vs1.RowSel, 1)
   PopUpValue6 = vs1.TextMatrix(vs1.RowSel, 1)
   pname_ = cboParty.text
   frmApproval.Show

End If

Screen.MousePointer = vbDefault

End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 116 Then
    Call cmdRefresh_Click
    cboParty.SetFocus
    Exit Sub
End If


Screen.MousePointer = vbHourglass
If KeyCode = 13 Then
If vs1.TextMatrix(vs1.RowSel, 0) = "I" Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         s1 = 1
         invoice.Show
ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "CI" Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         s1 = 12
         Critnote.Show
ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "CN" Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         Creditnotefile.Show
ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "DN" Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         Debitnotefile.Show
ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "C/M" Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         countersale.Show
End If


End If


If KeyCode = 112 Then
   txtdes.SetFocus
End If

Screen.MousePointer = vbDefault

End Sub
Private Sub vs1_SelChange()
 search
 'On Error Resume Next
 'If RS.State = 1 Then RS.close
 'RS.Open "select DNo,DDate,ScName,RepName from "
 'lblSt.Caption = vs1.TextMatrix(vs1.RowSel, 1)
End Sub
Private Sub vsop_Click()


If vsop.Col = 0 Then
   vsop.Editable = flexEDNone
ElseIf vsop.Col = 1 Then
   vsop.Editable = flexEDNone
ElseIf vsop.Col = 2 Then
   vsop.Editable = flexEDKbdMouse
ElseIf vsop.Col = 3 Then
   vsop.Editable = flexEDKbdMouse
End If
  

End Sub

Private Sub vsop_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    I = 1
    cboParty.text = vsop.TextMatrix(vsop.RowSel, 1)
    PopUpValue2 = cboParty.text
    vs1.Clear
    fillGrid
    SaveDatainTempledger
    CalculateTotalDrCr
    setWidth
    PopUpValue1 = ""
    Opening.Tab = 1
End If
End Sub

Private Sub vsop_SelChange()
col_name = vsop.TextMatrix(0, vsop.Col)

If col_name = "City" Then
   col_name = "DISTCODE"
ElseIf col_name = "Party" Then
   col_name = "SUBLEDGER"
ElseIf col_name = "Opening" Then
   col_name = "op"
End If

End Sub
Private Sub VsOrderList_DblClick()

If VsOrderList.TextMatrix(VsOrderList.RowSel, 0) <> "" Then
   popupvalue5 = VsOrderList.TextMatrix(VsOrderList.RowSel, 0)
   frmINVOrder.Show
   s1 = 100
End If

End Sub
Private Sub VsOrderList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
  
  If VsOrderList.TextMatrix(VsOrderList.RowSel, 0) <> "" Then
   popupvalue5 = VsOrderList.TextMatrix(VsOrderList.RowSel, 0)
   frmINVOrder.Show
   s1 = 100
  End If

  End If
End Sub
