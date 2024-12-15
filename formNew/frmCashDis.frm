VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCashDis 
   Caption         =   "Cash Discount List"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   9720
   Begin VB.Frame panel 
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
      Height          =   7065
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   9585
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   615
         Left            =   8700
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton CommandPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   615
         Left            =   7800
         Picture         =   "frmCashDis.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View"
         Height          =   615
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox txtTotalCOD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5700
         TabIndex        =   58
         Top             =   6720
         Width           =   1215
      End
      Begin VB.TextBox txtPaymentDay 
         Height          =   315
         Left            =   2820
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtCod 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FAEFC9&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4620
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FAEFC9&
         Caption         =   "&City Print "
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12540
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2340
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FAEFC9&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12540
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1860
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FAEFC9&
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12660
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4140
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FAEFC9&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1860
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FAEFC9&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3660
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FAEFC9&
         Caption         =   "&Modify"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton cmdprintalf 
         BackColor       =   &H00FAEFC9&
         Caption         =   "&Print Alphabet"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12540
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2820
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Frame f1 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   11700
         TabIndex        =   6
         Top             =   7560
         Visible         =   0   'False
         Width           =   3270
         Begin VB.OptionButton party 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Party Wise Dr/Cr Entry"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3450
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   45
            Width           =   3360
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Bill Aouthorized Option"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   3375
         End
         Begin VB.Image Image1 
            Height          =   675
            Left            =   240
            Stretch         =   -1  'True
            Top             =   75
            Width           =   10155
         End
      End
      Begin TabDlg.SSTab Opening 
         Height          =   4185
         Left            =   12000
         TabIndex        =   17
         Top             =   1620
         Visible         =   0   'False
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   7382
         _Version        =   393216
         Tab             =   1
         TabsPerRow      =   6
         TabHeight       =   697
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Tab 0"
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label20"
         Tab(0).Control(1)=   "Label19"
         Tab(0).Control(2)=   "Label18"
         Tab(0).Control(3)=   "Line1"
         Tab(0).Control(4)=   "Label15"
         Tab(0).Control(5)=   "Label13"
         Tab(0).Control(6)=   "Label12"
         Tab(0).Control(7)=   "Label8"
         Tab(0).Control(8)=   "Label6"
         Tab(0).Control(9)=   "Label5"
         Tab(0).Control(10)=   "Label4"
         Tab(0).Control(11)=   "Line2"
         Tab(0).Control(12)=   "Label9"
         Tab(0).Control(13)=   "RecDates"
         Tab(0).Control(14)=   "selectAll"
         Tab(0).Control(15)=   "cboPartyList"
         Tab(0).Control(16)=   "cmddewali"
         Tab(0).Control(17)=   "txtalfa"
         Tab(0).Control(18)=   "cboStation"
         Tab(0).Control(19)=   "closingcr"
         Tab(0).Control(20)=   "txtBalance"
         Tab(0).Control(21)=   "cmdShow1"
         Tab(0).Control(22)=   "cboop"
         Tab(0).Control(23)=   "txtOp"
         Tab(0).Control(24)=   "txtRecno"
         Tab(0).Control(25)=   "txtQty"
         Tab(0).Control(26)=   "Receive"
         Tab(0).Control(27)=   "Issue"
         Tab(0).Control(28)=   "txtdes"
         Tab(0).ControlCount=   29
         TabCaption(1)   =   "Dr/Cr Entry"
         TabPicture(1)   =   "frmCashDis.frx":0BE4
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label11"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "phone"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "CrLebel"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "drLebel"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "crpt"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtClosing"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Check1"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtcr"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txtRem"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "comdio"
         Tab(2).ControlCount=   1
         Begin VB.TextBox txtdes 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   -73710
            MaxLength       =   50
            TabIndex        =   36
            Top             =   1290
            Visible         =   0   'False
            Width           =   4095
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
            Left            =   -72270
            TabIndex        =   35
            Top             =   1830
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   675
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
            Left            =   -71640
            TabIndex        =   34
            Top             =   1830
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   -73710
            MaxLength       =   8
            TabIndex        =   33
            Top             =   1860
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtRecno 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   -73710
            TabIndex        =   32
            Top             =   630
            Visible         =   0   'False
            Width           =   1335
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
            Left            =   -66645
            MaxLength       =   50
            TabIndex        =   31
            Top             =   1080
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.ComboBox cboop 
            Height          =   315
            Left            =   -65265
            Style           =   1  'Simple Combo
            TabIndex        =   30
            Top             =   1080
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.CommandButton cmdShow1 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   -70260
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   600
            Visible         =   0   'False
            Width           =   225
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
            Left            =   -66645
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1560
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.ComboBox closingcr 
            Height          =   315
            Left            =   -65265
            Style           =   1  'Simple Combo
            TabIndex        =   27
            Top             =   1560
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.ComboBox cboStation 
            Height          =   315
            Left            =   -67905
            TabIndex        =   26
            Top             =   2280
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox txtalfa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   -64890
            MaxLength       =   8
            TabIndex        =   25
            Top             =   2295
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmddewali 
            Caption         =   "Refresh Diwali Head"
            Height          =   255
            Left            =   -66645
            TabIndex        =   24
            Top             =   600
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.ListBox cboPartyList 
            Appearance      =   0  'Flat
            Height          =   2730
            Left            =   -67920
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   2595
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.CheckBox selectAll 
            Caption         =   "Select All"
            Height          =   300
            Left            =   -65880
            TabIndex        =   22
            Top             =   2295
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox txtRem 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   1005
            Left            =   7980
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   660
            Visible         =   0   'False
            Width           =   255
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
            Left            =   8010
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   7755
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Search FA"
            Height          =   210
            Left            =   4050
            TabIndex        =   19
            Top             =   1920
            Visible         =   0   'False
            Width           =   90
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
            Left            =   6540
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   7755
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSComDlg.CommonDialog comdio 
            Left            =   -72840
            Top             =   6300
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Crystal.CrystalReport crpt 
            Left            =   11505
            Top             =   390
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
            WindowShowSearchBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin MSComCtl2.DTPicker RecDates 
            Height          =   285
            Left            =   -71610
            TabIndex        =   37
            Top             =   615
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   71892995
            UpDown          =   -1  'True
            CurrentDate     =   37701
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E98A0A&
            BackStyle       =   0  'Transparent
            Caption         =   "Esc To Exit"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -64620
            TabIndex        =   52
            Top             =   600
            Width           =   975
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   -74940
            X2              =   -63540
            Y1              =   2220
            Y2              =   2220
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
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
            Left            =   -74925
            TabIndex        =   51
            Top             =   1290
            Visible         =   0   'False
            Width           =   1335
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
            Left            =   -74910
            TabIndex        =   50
            Top             =   1875
            Visible         =   0   'False
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
            Height          =   255
            Left            =   -72240
            TabIndex        =   49
            Top             =   660
            Visible         =   0   'False
            Width           =   750
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
            Left            =   -74910
            TabIndex        =   48
            Top             =   630
            Visible         =   0   'False
            Width           =   975
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
            Height          =   270
            Left            =   -66645
            TabIndex        =   47
            Top             =   900
            Visible         =   0   'False
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
            Height          =   270
            Left            =   -66645
            TabIndex        =   46
            Top             =   1380
            Visible         =   0   'False
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
            Height          =   270
            Left            =   -66660
            TabIndex        =   45
            Top             =   1920
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   -74940
            X2              =   -63540
            Y1              =   540
            Y2              =   540
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
            Left            =   -74895
            TabIndex        =   44
            Top             =   1605
            Visible         =   0   'False
            Width           =   1335
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
            Height          =   195
            Left            =   -73740
            TabIndex        =   43
            Top             =   360
            Visible         =   0   'False
            Width           =   2955
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
            Height          =   195
            Left            =   -70860
            TabIndex        =   42
            Top             =   1860
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label drLebel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4980
            TabIndex        =   41
            Top             =   6960
            Width           =   1335
         End
         Begin VB.Label CrLebel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   6300
            TabIndex        =   40
            Top             =   6960
            Width           =   1275
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
            Left            =   1290
            TabIndex        =   39
            Top             =   1695
            Width           =   6975
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Total "
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   5745
            TabIndex        =   38
            Top             =   7815
            Visible         =   0   'False
            Width           =   795
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   5760
         Left            =   60
         TabIndex        =   53
         Top             =   900
         Width           =   9420
         _cx             =   16616
         _cy             =   10160
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Rows            =   50
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin MSComCtl2.DTPicker todate 
         Height          =   375
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         _Version        =   393216
         Format          =   71892993
         CurrentDate     =   39100
      End
      Begin MSComCtl2.DTPicker fromdate 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         _Version        =   393216
         Format          =   71892993
         CurrentDate     =   39100
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Total  Amt."
         Height          =   255
         Left            =   4620
         TabIndex        =   59
         Top             =   6780
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   255
         Left            =   3600
         TabIndex        =   57
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   255
         Left            =   5340
         TabIndex        =   56
         Top             =   300
         Width           =   285
      End
      Begin VB.Label lblAmount_days 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment with in (days)"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   1500
         TabIndex        =   55
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label percentage 
         BackStyle       =   0  'Transparent
         Caption         =   "CD(%) "
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCashDis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdView_Click()

Dim date1, date2
Dim cdAmt As Double
cdAmt = 0

'''Dim a_strResult
'''a_strResult = Split(bill, ",")
'''If a_strResult > 0 Then
'''    For J = 0 To UBound(a_strResult)
'''    Next
'''Else
'''End If


vs.Rows = 2
I = 1

If RS.State = 1 Then RS.close
RS.Open "select  INVOICEDATE,SUBLEDGER,Amount,Category from CD_List where (INVOICEDATE>=convert(smalldatetime,'" & fromdate.value & "',103)  and INVOICEDATE<=convert(smalldatetime,'" & todate.value & "',103) ) and " & stringyear & " order by INVOICEDATE"
While RS.EOF = False


vs.TextMatrix(I, 0) = RS!invoicedate
vs.TextMatrix(I, 1) = RS!SUBLEDGER
vs.TextMatrix(I, 2) = RS!amount
vs.TextMatrix(I, 3) = 0

If RS!category = "Due" Then
    vs.TextMatrix(I, 4) = "Bill Amt."
    date1 = RS!invoicedate
Else
   vs.TextMatrix(I, 4) = "Payment Rec."
   date2 = RS!invoicedate
   d1 = DateDiff("d", date1, date2)
   If Val(txtPaymentDay.Text) >= d1 Then
        If Val(txtCod) > 0 Then
            vs.TextMatrix(I, 3) = Round((RS!amount * txtCod / 100), 0)
            cdAmt = cdAmt + Val(vs.TextMatrix(I, 3))
        End If
   End If
End If

I = I + 1
vs.Rows = vs.Rows + 1

RS.MoveNext
Wend

vs.FormatString = "DATE|PARTY NAME|AMOUNT|CD AMOUNT|Bill Amt./Payment Rec."


vs.ColWidth(0) = 1200
vs.ColWidth(1) = 4500
vs.ColWidth(2) = 1600
vs.ColWidth(3) = 1600
vs.ColWidth(4) = 1600

'------------------------------------------------------------
txtTotalCOD.Text = cdAmt


End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

Me.Width = 10000
Me.Height = 7800

fromdate.value = Format(fromDate_setup, "dd/MM/yyyy")
todate.value = Format(toDate_setup, "dd/MM/yyyy")

fillGrid
BackColorFrom Me, 1

End Sub
Sub fillGrid()
   vs.FormatString = "S.No.|Party Name |>COD Amount"
   vs.ColWidth(0) = 800
   vs.ColWidth(1) = 6000
   vs.ColWidth(2) = 2000
End Sub
