VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmShowLedger 
   ClientHeight    =   9156
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14988
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9156
   ScaleWidth      =   14988
   Begin VB.Frame panel 
      Caption         =   "View Ledger"
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
      Height          =   8925
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   14820
      Begin VB.CommandButton cmdPrintAgentLed 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   660
         Left            =   11676
         Picture         =   "frmShowLedger.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   276
         Width           =   1170
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         Height          =   660
         Left            =   12936
         Picture         =   "frmShowLedger.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   276
         Width           =   1125
      End
      Begin VB.Frame f1 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   14868
         TabIndex        =   47
         Top             =   7560
         Visible         =   0   'False
         Width           =   3270
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
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   120
            Width           =   3375
         End
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
            TabIndex        =   48
            Top             =   45
            Width           =   3360
         End
         Begin VB.Image Image1 
            Height          =   675
            Left            =   240
            Stretch         =   -1  'True
            Top             =   75
            Width           =   10155
         End
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
         Height          =   435
         Left            =   14916
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2820
         Visible         =   0   'False
         Width           =   870
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
         Height          =   435
         Left            =   14976
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3240
         Visible         =   0   'False
         Width           =   870
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
         Height          =   435
         Left            =   14940
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3660
         Visible         =   0   'False
         Width           =   840
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
         Height          =   435
         Left            =   14976
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1860
         Visible         =   0   'False
         Width           =   840
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
         Height          =   435
         Left            =   15000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4140
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmdPrint 
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
         Height          =   435
         Left            =   14880
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1860
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FAEFC9&
         Caption         =   "&City Print "
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   14916
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2340
         Visible         =   0   'False
         Width           =   825
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
         Height          =   435
         Left            =   14832
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4620
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboParty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   312
         Left            =   1320
         Style           =   1  'Simple Combo
         TabIndex        =   38
         Top             =   540
         Width           =   6135
      End
      Begin TabDlg.SSTab Opening 
         Height          =   4188
         Left            =   14880
         TabIndex        =   2
         Top             =   1620
         Visible         =   0   'False
         Width           =   3696
         _ExtentX        =   6519
         _ExtentY        =   7387
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
         TabPicture(1)   =   "frmShowLedger.frx":17C8
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "drLebel"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "CrLebel"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "phone"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label11"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "crpt"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtRem"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "txtcr"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Check1"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txtClosing"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "comdio"
         Tab(2).ControlCount=   1
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
            TabIndex        =   21
            Top             =   7755
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Search FA"
            Height          =   210
            Left            =   4050
            TabIndex        =   20
            Top             =   1920
            Visible         =   0   'False
            Width           =   90
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
            TabIndex        =   19
            Top             =   7755
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtRem 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00FF8080&
            Height          =   1005
            Left            =   7980
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   660
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox selectAll 
            Caption         =   "Select All"
            Height          =   300
            Left            =   -65880
            TabIndex        =   17
            Top             =   2295
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.ListBox cboPartyList 
            Appearance      =   0  'Flat
            Height          =   2616
            Left            =   -67920
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   2595
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.CommandButton cmddewali 
            Caption         =   "Refresh Diwali Head"
            Height          =   255
            Left            =   -66645
            TabIndex        =   15
            Top             =   600
            Visible         =   0   'False
            Width           =   1875
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
            Height          =   285
            Left            =   -64890
            MaxLength       =   8
            TabIndex        =   14
            Top             =   2295
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox cboStation 
            Height          =   315
            Left            =   -67905
            TabIndex        =   13
            Top             =   2280
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.ComboBox closingcr 
            Height          =   288
            Left            =   -65265
            Style           =   1  'Simple Combo
            TabIndex        =   12
            Top             =   1560
            Visible         =   0   'False
            Width           =   465
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
            TabIndex        =   11
            Top             =   1560
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CommandButton cmdShow1 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   -70260
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   600
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.ComboBox cboop 
            Height          =   288
            Left            =   -65265
            Style           =   1  'Simple Combo
            TabIndex        =   9
            Top             =   1080
            Visible         =   0   'False
            Width           =   465
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
            TabIndex        =   8
            Top             =   1080
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox txtRecno 
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
            Left            =   -73710
            TabIndex        =   7
            Top             =   630
            Visible         =   0   'False
            Width           =   1335
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
            Left            =   -73710
            MaxLength       =   8
            TabIndex        =   6
            Top             =   1860
            Visible         =   0   'False
            Width           =   1395
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
            TabIndex        =   5
            Top             =   1830
            Visible         =   0   'False
            Width           =   690
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
            TabIndex        =   4
            Top             =   1830
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   675
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
            Left            =   -73710
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1290
            Visible         =   0   'False
            Width           =   4095
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
            _ExtentX        =   593
            _ExtentY        =   593
            _Version        =   348160
            PrintFileLinesPerPage=   60
            WindowShowSearchBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin MSComCtl2.DTPicker RecDates 
            Height          =   285
            Left            =   -71610
            TabIndex        =   22
            Top             =   615
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2286
            _ExtentY        =   508
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   141557763
            UpDown          =   -1  'True
            CurrentDate     =   37701
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
            Left            =   5745
            TabIndex        =   37
            Top             =   7815
            Visible         =   0   'False
            Width           =   795
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
            TabIndex        =   36
            Top             =   1695
            Width           =   6975
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
            Left            =   6300
            TabIndex        =   35
            Top             =   6960
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
            Left            =   4980
            TabIndex        =   34
            Top             =   6960
            Width           =   1335
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
            TabIndex        =   33
            Top             =   1860
            Visible         =   0   'False
            Width           =   2715
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
            TabIndex        =   32
            Top             =   360
            Visible         =   0   'False
            Width           =   2955
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
            TabIndex        =   31
            Top             =   1605
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   -74940
            X2              =   -63540
            Y1              =   540
            Y2              =   540
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
            TabIndex        =   30
            Top             =   1920
            Visible         =   0   'False
            Width           =   840
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
            TabIndex        =   29
            Top             =   1380
            Visible         =   0   'False
            Width           =   780
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
            TabIndex        =   28
            Top             =   900
            Visible         =   0   'False
            Width           =   960
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
            TabIndex        =   27
            Top             =   630
            Visible         =   0   'False
            Width           =   975
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
            TabIndex        =   26
            Top             =   660
            Visible         =   0   'False
            Width           =   750
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
            TabIndex        =   25
            Top             =   1875
            Visible         =   0   'False
            Width           =   1095
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
            TabIndex        =   24
            Top             =   1290
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   -74940
            X2              =   -63540
            Y1              =   2220
            Y2              =   2220
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
            Height          =   255
            Left            =   -64620
            TabIndex        =   23
            Top             =   600
            Width           =   975
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid vs1 
         Height          =   6948
         Left            =   120
         TabIndex        =   53
         Top             =   1836
         Width           =   14616
         _cx             =   25781
         _cy             =   12255
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
         BackColor       =   16448711
         ForeColor       =   16711680
         BackColorFixed  =   16251308
         ForeColorFixed  =   255
         BackColorSel    =   16448755
         ForeColorSel    =   16744448
         BackColorBkg    =   16251308
         BackColorAlternate=   16448711
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   450
         RowHeightMax    =   0
         ColWidthMin     =   700
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
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   768
         Left            =   108
         TabIndex        =   56
         Top             =   1008
         Width           =   7332
         _cx             =   12933
         _cy             =   1355
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   11162880
         BackColorFixed  =   12648447
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   -2147483647
         BackColorBkg    =   -2147483636
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmShowLedger.frx":17E4
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
         Editable        =   1
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 For  Search Ship To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   4560
         TabIndex        =   55
         Top             =   300
         Width           =   3108
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 For  Search Representative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   1644
         TabIndex        =   54
         Top             =   300
         Width           =   2652
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name "
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   50
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblId 
         Height          =   330
         Left            =   4500
         TabIndex        =   1
         Top             =   630
         Visible         =   0   'False
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmShowLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim bb As Boolean

Dim bb2 As Boolean
Dim rss As New ADODB.Recordset
Dim from_date As Date
Dim I As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim search_v As Boolean
Dim to_date As Date
Dim kk As Integer
Dim bb1 As Boolean
Dim str1 As New ADODB.Recordset
Dim din_ As Boolean

Sub vsIni()

   
End Sub

Private Sub All_Click()
If All.value = True Then
'    Call cmdShow_Click
End If

End Sub

Private Sub autho_Click()
If autho.value = True Then
'    Call cmdShow_Click
End If
End Sub



Private Sub cash_Click()
    If cash.value = True Then
'       Call cmdShow_Click
    End If
End Sub

Private Sub cboop_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtdes.SetFocus
   End If
End Sub

Private Sub cboStation_Click()
cboPartyList.Visible = True
If RS.State = 1 Then RS.close
RS.Open "select distinct(AgentName) from SLEDGER where " & stringyear & " and DISTCODE='" & cboStation.Text & "'", con
cboPartyList.Clear
While RS.EOF = False
cboPartyList.AddItem RS(0)
RS.MoveNext
Wend
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

Private Sub cmdAson_Click()
'showDataAsOn dateason
End Sub

Private Sub cmddewali_Click()
    Dim f As New ADODB.Recordset
    If f.State = 1 Then f.close
    f.Open "select AMOUNT,text,INVOICENO from invoicec where " & stringyear & " and TEXT = '" & "DIWALI SPECIAL" & "' and AMOUNT>0", con
    While f.EOF = False
        con.Execute "update INVOICEA_sp set t2='" & f.Fields("amount").value & "' where " & stringyear & " and INVOICENO=" & f.Fields("INVOICENO").value & ""
        f.MoveNext
    Wend
    If f.State = 1 Then f.close
    f.Open "select AMOUNT,text,INVOICENO from CASHC where " & stringyear & " and TEXT = '" & "DIWALI SPECIAL" & "' and AMOUNT>0", con
    While f.EOF = False
        con.Execute "update CASHA set t2='" & f.Fields("amount").value & "' where " & stringyear & " and INVOICENO=" & f.Fields("INVOICENO").value & ""
        f.MoveNext
    Wend
    MsgBox "Data Refresh...", vbInformation
End Sub

Private Sub cmdPath_Click()
Me.comdio.ShowOpen
'Me.txtPath.Text = Me.comdio.FileName
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()

On Error GoTo aa10
Screen.MousePointer = vbHourglass
Dim op, drcr
Dim rs1 As New ADODB.Recordset
con.Execute "delete from templedger1"

If RS.State = 1 Then RS.close
RS.Open "select AgentName from SLEDGER where " & stringyear & " and AgentName = '" + Trim(cboParty.Text) + "'", con

While RS.EOF = False

'==Code For Opening=============================================
con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype)  SELECT op,drcr,AgentName,'Opening' from sledger where " & stringyear & " and AgentName = '" + RS.Fields(0).value + "'   group by op,AgentName,drcr HAVING  op <> 0;"
If rs1.State = 1 Then rs1.close
rs1.Open "SELECT op,drcr from sledger where " & stringyear & " and AgentName = '" + RS.Fields(0).value + "'", con
If Not IsNull(rs1.Fields(0).value) Then
   op = Val(rs1.Fields(0).value)
   drcr = rs1.Fields(1).value
Else
   op = 0
End If

'==============================================
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,AgentName from INVOICEA_sp where " & stringyear & " and AgentName='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',BAA,netamount,AgentName from invoicea_spRet where " & stringyear & " and AgentName='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,AgentName from CASHA where  " & stringyear & " and AgentName='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa where " & stringyear & " and  psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a where " & stringyear & " and psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty where " & stringyear & " and PartyName='" & RS.Fields(0).value & "' order by dates,recno"
'===============================================================
If op <> 0 Then
con.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "' where " & stringyear & " and party = '" + RS.Fields(0).value + "' and Billtype<>'Opening'"
End If
'===============================================================
Sleep (200)
RS.MoveNext
Wend

DSNNew

Sleep (300)
crpt.Reset
'crpt.ReportFileName = App.Path & "\" & directory & "\PartyLedger.rpt"
crpt.ReportFileName = st1 & "\" & directory & "\PartyLedger.rpt"
crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"
crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.Action = 1
Screen.MousePointer = vbDefault
Exit Sub
aa10:
MsgBox err.DESCRIPTION
End Sub
Private Sub cmdPrint1_Click()

crpt.Reset

If Check_ClosingDesc.value = 1 Then
   crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseClosing_descClosing.rpt"
Else
   crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseClosing.rpt"
End If

crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"

''======================================================================
''======================================================================

If Check2.value = 0 Then

    If cboStation1.Text <> "" And txtAmount.Text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.Text <> "" And txtAmount.Text <> "" Then
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.DISTCODE}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.Text & ""
    End If
    
    
    ElseIf cboStation1.Text = "" And txtAmount.Text <> "" Then
    
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.Text & ""
    End If
    
    Else
    crpt.ReplaceSelectionFormula "abs({tempLedgerRpt.Owner})<>0"
    End If


ElseIf Check2.value = 1 Then


    If cboStation1.Text <> "" And txtAmount.Text = "" Then
    crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0"
    ElseIf cboStation1.Text <> "" And txtAmount.Text <> "" Then
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.states}='" & cboStation1.Text & "' and {tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.Text & ""
    End If
    
    
    ElseIf cboStation1.Text = "" And txtAmount.Text <> "" Then
    
    If Val(txtAmount) >= 0 Then
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}>=" & txtAmount.Text & ""
    Else
       crpt.ReplaceSelectionFormula "{tempLedgerRpt.Owner}<>0 and {tempLedgerRpt.Owner}<=" & txtAmount.Text & ""
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
crpt.Formulas(0) = "partyname='" & cboStation1.Text & "'"
crpt.Formulas(1) = "ason='" & dateAson.value & "'"

crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.WindowShowSearchBtn = True
crpt.Action = 1

End Sub

Private Sub cmdPrintAgentLed_Click()

Screen.MousePointer = vbHourglass
DSNNew

With crpt
 .Reset
 .ReportFileName = rptPath & "/AgentLedger.rpt"
 .Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 .ReplaceSelectionFormula "{tempLedgerRpt.party}='" & cboParty.Text & "'"
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .Action = 1
End With
Screen.MousePointer = vbDefault


End Sub
Private Sub cmdprintalf_Click()
 
 If txtalfa.Text = "" Then
    MsgBox "Please Enter Alphabet...", vbInformation
    Exit Sub
 End If
 
 Screen.MousePointer = vbHourglass
 CityWiseStatement
 Screen.MousePointer = vbDefault

End Sub

Private Sub cmdset_Click()
   
If RS.State = 1 Then RS.close
RS.Open "select * from pass where pass='" & cp & "'", con
If RS.EOF = True Then
MsgBox "Enter Valid Password !!", vbInformation
Exit Sub
End If
   
    
saveData
   
End Sub
Sub saveData()
   
''''   If MsgBox("Want to Set ?", vbQuestion + vbYesNo) = vbYes Then
''''
''''   Screen.MousePointer = vbHourglass
''''   'cmdShow1.Visible = True
''''
''''
''''   If sales.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''             CON.Execute "update INVOICEA_sp set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''             CON.Execute "update INVOICEA_sp set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''  ElseIf credit.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''            CON.Execute "update invoicea_spRet set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and  INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''            CON.Execute "update invoicea_spRet set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''
''''  ElseIf cash.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''            CON.Execute "update CASHA set BAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''            CON.Execute "update CASHA set BAuthorized=" & vs.TextMatrix(J, 5) & " where INVOICENO=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''  ElseIf crdit.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''            CON.Execute "update cnf1a set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and cnn=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''            CON.Execute "update cnf1a set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and cnn=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''
''''  ElseIf dbit.value = True Then
''''
''''        For J = 1 To vs.Rows - 1
''''          If vs.TextMatrix(J, 5) = True Then
''''            CON.Execute "update dnfa set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and dnn=" & vs.TextMatrix(J, 1) & ""
''''            Else
''''            CON.Execute "update dnfa set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and dnn=" & vs.TextMatrix(J, 1) & ""
''''          End If
''''        Next
''''
''''
''''  End If
''''
''''
''''   End If
''''
''''
'''' Screen.MousePointer = vbDefault
End Sub
Sub SearchFa()
      
      If RS.State = 1 Then RS.close
      If din_ = False Then
         RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,BAA,t2,ScName from INVOICEA_sp where " & stringyear & " and AgentName='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
      Else
         RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,BAA,t2,ScName from INVOICEA_sp where " & stringyear & " and shipto='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
      End If
      
      If RS.EOF = False Then
        vs1.rows = (vs1.rows + RS.RecordCount)
        For I = I To vs1.rows - 1
        If RS.EOF = False Then
           vs1.TextMatrix(I, 0) = "I"
           vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
           vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
           If IsNull(RS.Fields("t2").value) Then
              vs1.TextMatrix(I, 3) = "Issue"
           Else
              vs1.TextMatrix(I, 3) = "Invoice Sales" & RS.Fields("t2").value & " " & "DS"
           End If
           vs1.TextMatrix(I, 4) = Format(RS.Fields("netamount").value, "0.00")
           vs1.TextMatrix(I, 5) = Format(RS.Fields("BAA").value, "0.00")
           
           If RS.Fields("INVOICENO").value = 3009 Then
           '   V = 1
           End If
           
           vs1.TextMatrix(I, 8) = RS.Fields("scname").value
           
            RS.MoveNext
         End If
        Next
      End If
      
      
      
    '================
     If RS.State = 1 Then RS.close
     RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,baa from invoicea_spRet where " & stringyear & " and AgentName='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
     'RS.Open "select INVOICENO,INVOICEDATE,AgentName,netamount,baa from invoicea_spRet where " & stringyear & " and AgentName='" & cboParty.Text & "'", CON
     If RS.EOF = False Then
        vs1.rows = vs1.rows + RS.RecordCount
        For I = I To vs1.rows - 1
         
        If RS.EOF = False Then
         vs1.TextMatrix(I, 0) = "R"
         vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
         vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
         vs1.TextMatrix(I, 3) = "Return"
         vs1.TextMatrix(I, 4) = Format(RS.Fields("baa").value, "0.00")
         vs1.TextMatrix(I, 5) = Format(RS.Fields("netamount").value, "0.00")
         'vs1.TextMatrix(I, 8) = RS.Fields("scname").value
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
       DSNNew
       
       s = ""
       Dim rs1 As New ADODB.Recordset
       con.Execute "delete from templedger1 " & stringyear & ""
       If RS.State = 1 Then RS.close
       If cboStation.Text <> "" And txtalfa.Text = "" Then
       For I = 0 To cboPartyList.ListCount - 1
        If cboPartyList.Selected(I) = True Then
        If s = "" Then
          s = "AgentName " & " = " & "'" & cboPartyList.List(I) & "'"
        Else
          s = s & " or " & "AgentName " & " = " & "'" & cboPartyList.List(I) & "'"
        End If
        End If
       Next
       
       If s = "" Then
        If RS.State = 1 Then RS.close
        RS.Open "select AgentName from SLEDGER where " & stringyear & " and DISTCODE = '" & cboStation.Text & "'", con
       Else
        If RS.State = 1 Then RS.close
        RS.Open "select AgentName from SLEDGER where " & stringyear & " and " & s, con
       End If
       
       ElseIf txtalfa.Text <> "" And cboStation.Text = "" Then
        RS.Open "select AgentName from SLEDGER where " & stringyear & " and AgentName like '" + Trim(txtalfa.Text) + "%'", con
       Else
         Exit Sub
       End If
       While RS.EOF = False
           '==Code For Opening=============================================
            con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype)  SELECT op,drcr,AgentName,'Opening' from sledger where " & stringyear & " and AgentName = '" + RS.Fields(0).value + "'   group by op,AgentName,drcr HAVING  op <> 0;"
            If rs1.State = 1 Then rs1.close
            rs1.Open "SELECT op,drcr from sledger where " & stringyear & " and AgentName = '" + RS.Fields(0).value + "'", con
            If Not IsNull(rs1.Fields(0).value) Then
               op = Val(rs1.Fields(0).value)
               drcr = rs1.Fields(1).value
            Else
               op = 0
            End If
           '==============================================
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,AgentName from INVOICEA_sp where " & stringyear & " and AgentName='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,AgentName from invoicea_spRet where " & stringyear & " and AgentName='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,AgentName from CASHA where  " & stringyear & " and AgentName='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party) Select cnd,'CN',cnn,'Credit Note',na,'0',psld from Cnf1a where " & stringyear & " and  psld='" & RS.Fields(0).value & "'"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party) Select dates,'J',Recno,Particullar,Dr,CR,PartyName from ReceiveIssueParty where " & stringyear & " and PartyName='" & RS.Fields(0).value & "' order by dates,recno"
          '===============================================================
          If op <> 0 Then
           con.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "' where " & stringyear & " and party = '" + RS.Fields(0).value + "' and Billtype<>'Opening'"
          End If
          '===============================================================
           RS.MoveNext
       Wend
       DoEvents
       MsgBox "View"
 crpt.Reset
 'crpt.ReportFileName = App.Path & "\" & directory & "\PartyLedger.rpt"
 crpt.ReportFileName = st1 & "\" & directory & "\PartyLedger.rpt"
 crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"
 crpt.WindowShowPrintSetupBtn = True
 crpt.Formulas(0) = "partyname='" & cboStation.Text & "'"
 crpt.WindowShowPrintBtn = True
 crpt.WindowState = crptMaximized
 crpt.Action = 1
End Sub
Private Sub cmdShowClosing_Click()

bb1 = False
'showData

End Sub

Private Sub cmdupdatep_Click()
   Dim partyname
   Dim pcode
   partyname = ""
   pcode = ""
   
    
   If RS.State = 1 Then RS.close
   RS.Open "select AgentName from sledger", con
   While RS.EOF = False
       
       aa = InStr(RS(0), " ")
       partyname = Mid(RS(0), aa)
       pcode = Mid(RS(0), 1, aa)
       
       con.Execute "update  Sledger  set party='" & Trim(partyname) & "',code='" & Trim(pcode) & "' where " & stringyear & " and AgentName='" & RS(0) & "'"
       
       RS.MoveNext
       
   Wend
   
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
          con.Execute "update SLEDGER set op=" & CDbl(vsop.TextMatrix(I, 2)) & ",drcr='" & vsop.TextMatrix(I, 3) & "' where " & stringyear & " and AgentName='" & vsop.TextMatrix(I, 1) & "'"
       End If
   Next
   
   Screen.MousePointer = vbDefault
   

   
End If
   

   
   
   
   
   
End Sub
Private Sub Command2_Click()

DSNNew

crpt.Reset
crpt.ReportFileName = st1 & "\" & directory & "\PartyWiseDrClosing.rpt"
crpt.DataFiles(0) = st1 & "\" + directory + "\data.mdb"
crpt.ReplaceSelectionFormula "{tempLedgerRpt.Offdays}='" & "1" & "' and {tempLedgerRpt.Owner}>=" & 1 & ""
DoEvents
MsgBox ("View")
crpt.Formulas(0) = "partyname='" & cboStation1.Text & "'"
crpt.WindowShowPrintSetupBtn = True
crpt.WindowShowPrintBtn = True
crpt.WindowState = crptMaximized
crpt.WindowShowSearchBtn = True
crpt.Action = 1

End Sub

Private Sub Command3_Click()
 
 If cboStation.Text = "" Then
    MsgBox "Please Select Station...", vbInformation
    Exit Sub
 End If
 
 Screen.MousePointer = vbHourglass
 CityWiseStatement
 cboPartyList.Visible = False
 Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()

Dim FSO As FileSystemObject
Dim f As File
Dim txt As TextStream
Dim matter As String
Dim Total As String
Dim s(1, 2) As String
Set FSO = New FileSystemObject
Dim ss As String
'
Dim s1

matter = ""

Set txt = FSO.CreateTextFile(App.Path & "\mobile.txt", True)

If RS.State = 1 Then RS.close
If Check2.value = 0 Then
RS.Open "select mobile from sledger where " & stringyear & " and distcode='" & cboStation1.Text & "'", con, adOpenKeyset, adLockReadOnly
Else
RS.Open "select mobile from sledger where " & stringyear & " and states='" & cboStation1.Text & "'", con, adOpenKeyset, adLockReadOnly
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

Private Sub crdit_Click()
    If crdit.value = True Then
'       Call cmdShow_Click
    End If
End Sub

Private Sub credit_Click()
    If credit.value = True Then
'       Call cmdShow_Click
    End If

End Sub

Private Sub dbit_Click()
   If dbit.value = True Then
'       Call cmdShow_Click
    End If
End Sub

Private Sub Form_Activate()
' Me.WindowState = 2
  
End Sub
Sub searchAllotment(stt As String)

On Error GoTo err:

Dim rss As New ADODB.Recordset
Screen.MousePointer = vbHourglass
vs.Cols = 6
vs.rows = 2

Dim tqty As Long
Dim Ordqty As Long
Dim Spqty As Long

Dim fdate As String
Dim tdate As String

k1 = 1
vs.Clear

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT * from SpAllotmentQty where RepName='" & stt & "'", con
If rs1.EOF = False Then
   
   
    fdate = rs1!FromDate
    tdate = rs1!toDate
   
    For J = 1 To rs1.RecordCount
    
        DoEvents
        vs.TextMatrix(k1, 0) = k1
        vs.TextMatrix(k1, 1) = rs1!RepName
        vs.TextMatrix(k1, 2) = rs1!qty
        
        'If rs1!RepName = "SHIVRATAN RAWAT" Then
        'g = 0
        'End If
        
        If rss.State = 1 Then rss.close
        rss.Open "select sum(Qty) from totalSpQty_Issued where RepName='" & rs1!RepName & "'and (convert(datetime,invoiceDate,103)>=convert(datetime,'" & fdate & "',103) and convert(datetime,invoiceDate,103)<=convert(datetime,'" & tdate & "',103))", con
        If Not IsNull(rss(0)) Then
           vs.TextMatrix(k1, 3) = rss(0)
        Else
           vs.TextMatrix(k1, 3) = 0
        End If
        
                
        If rss.State = 1 Then rss.close
        rss.Open "SELECT sum(QTY) FROM TotalSpReturnRepWise where agentname='" & rs1!RepName & "'  and (convert(datetime,invoiceDate,103)>=convert(datetime,'" & fdate & "',103) and convert(datetime,invoiceDate,103)<=convert(datetime,'" & tdate & "',103))", con
        If Not IsNull(rss(0)) Then
           vs.TextMatrix(k1, 4) = rss(0)
        Else
           vs.TextMatrix(k1, 4) = 0
        End If
        
        
        tqty = IIf(vs.TextMatrix(k1, 2) = "", 0, vs.TextMatrix(k1, 2))
        Ordqty = IIf(vs.TextMatrix(k1, 3) = "", 0, vs.TextMatrix(k1, 3))
        Spqty = IIf(vs.TextMatrix(k1, 4) = "", 0, vs.TextMatrix(k1, 4))
        
        
        vs.TextMatrix(k1, 5) = ((tqty + Spqty) - Ordqty)
        
        For k2 = 0 To 5
          If Val(vs.TextMatrix(k1, 5)) < 0 Then
            vs.Cell(flexcpBackColor, k1, k2) = vbGreen
            DoEvents
          End If
        Next
        
        
        rs1.MoveNext
        k1 = k1 + 1
        'vs.rows = vs.rows + 1
        DoEvents
        DoEvents
    
    Next



End If

Screen.MousePointer = vbDefault

vs.FormatString = "||Allotment Qty.|TQty.Specimen|TQtyRet.Specimen|BalanceQty"
vs.ColWidth(0) = 0
vs.ColWidth(1) = 0
vs.ColWidth(2) = 1400
vs.ColWidth(3) = 1500
vs.ColWidth(4) = 1700
vs.ColWidth(5) = 1600

Exit Sub
err:
MsgBox "" & err.DESCRIPTION

Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Me.Top = 100
Me.Left = 100

Me.Width = 14900
Me.Height = 10100



vsIni
On Error Resume Next

kk = 1
'setwidth
'AddParty

dateAson.value = Date




FromDate.value = Date
toDate.value = Date
from_date = FromDate.value
'FillGrid

maxId
setWidth

cboop.ListIndex = 0



vsop.Cols = 6
vsop.TextMatrix(0, 0) = "City"
vsop.TextMatrix(0, 1) = "Party"
vsop.TextMatrix(0, 2) = "Opening"
vsop.TextMatrix(0, 3) = "Dr/Cr"
vsop.TextMatrix(0, 4) = "Closing Balance"
vsop.TextMatrix(0, 5) = "Dr/Cr"


vsop.ColWidth(0) = 1800
vsop.ColWidth(1) = 3200
vsop.ColWidth(2) = 1400
vsop.ColWidth(3) = 500
vsop.ColWidth(4) = 1400
vsop.ColWidth(5) = 500


If RS.State = 1 Then RS.close
RS.Open "select yarfrom,yarto from setup1 where " & stringyear & "", con
If RS.EOF = False Then
   FromDate.value = RS.Fields(0).value
   
   If (DateValue(RS!yarfrom) <= DateValue(Date) And DateValue(RS!yarto) >= DateValue(Date)) Then
      RecDates.value = Date
   Else
      RecDates.value = RS.Fields(1).value
   End If
   
End If

Me.Top = 50
Me.Left = 50


Opening.Tab = 1


If RS.State = 1 Then RS.close
RS.Open "select * from setup1 where " & stringyear & "", con, adOpenDynamic, adLockReadOnly, adCmdTable
If RS.EOF = False Then
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
End If



bb1 = False


fetchTab2

BackColorFrom Me

Screen.MousePointer = vbDefault

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
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub Opening_Click(PreviousTab As Integer)
      
     ' Screen.MousePointer = vbHourglass
      
      
      Dim closing As Double
      
      
      closing = 0
      
      If Opening.Tab = 0 Then
         
'         Call cmdShow_Click
         
      ElseIf Opening.Tab = 2 Then
       
        
'
      
      
      End If
      
      
      
      'Screen.MousePointer = vbDefault
      
End Sub
Sub fetchTab2()

        Screen.MousePointer = vbHourglass

        Dim fillVs As New ADODB.Recordset
        If fillVs.State = 1 Then fillVs.close
        'fillvs.Open "select DISTCODE as City,AgentName as Party,op,drcr from closing where gledger='SUNDRY DEBTORS'", con
        fillVs.Open "SELECT SLEDGER.DISTCODE,SLEDGER.AgentName,SLEDGER.OP,SLEDGER.drcr,(Sum(templedger1.Dr)-Sum(templedger1.Cr)) AS bal1 FROM SLEDGER LEFT JOIN templedger1 ON SLEDGER.AgentName = templedger1.Party where " & stringyear & " and  gledger='SUNDRY DEBTORS' GROUP BY SLEDGER.AgentName,SLEDGER.DISTCODE,[SLEDGER.OP], SLEDGER.drcr, SLEDGER.gledger", con

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

Private Sub party_Click()
   
   If party.value = True Then
      bill.Visible = False
      frmReceiveFromParty.Show
   Else
      partydrcr.Visible = False
      bill.Visible = True
   End If
   
   frmReceiveFromParty.Top = 800

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

Private Sub sales_Click()
    If sales.value = True Then
'       Call cmdShow_Click
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

Private Sub txtadmin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   setsecurity
   'pass.Visible = False
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

Private Sub txtdes_LostFocus()
    txtdes.BackColor = &HFFFFFF
End Sub

Private Sub txtOp_GotFocus()
txtOp.BackColor = &HFFFFC0
End Sub
Private Sub txtParty_GotFocus()
   If PopUpValue1 <> "" Then
      txtParty.Text = PopUpValue1
   End If
End Sub
Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 113 Then
       value = "select AgentName from INVOICEA_sp where " & stringyear & "  order by AgentName"
       popuplistModel10 value, con
    End If
End Sub
Private Sub txtParty_LostFocus()
  PopUpValue1 = ""
End Sub
Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
   If Val(txtQty.Text) = 0 Then
      txtQty.SetFocus
      Exit Sub
   End If
   If cmdSave.Enabled = True Then
      cmdSave.SetFocus
   End If
   End If
End Sub
Private Sub txtRem_LostFocus()
  If cboParty.Text <> "" Then
  If MsgBox("Want To Change Remarks ?", vbQuestion + vbYesNo) = vbYes Then
      con.Execute "update sledger set PartyRemarks = '" & txtRem.Text & "' where " & stringyear & " and AgentName='" & cboParty.Text & "'"
     
  End If
  End If
End Sub
Private Sub Unautho_Click()
If Unautho.value = True Then
'    Call cmdShow_Click
End If
End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)



Screen.MousePointer = vbHourglass
If KeyCode = 13 Then
If vs1.TextMatrix(vs1.RowSel, 0) = "I" Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         frmBookIssueSp.Show
ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "R" Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         frmBookIssueSp_Ret.Show
'ElseIf credit.value = True Then
'   If vs1.Col = 1 Then
'         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
'         'MainMenu.Toolbar1.Visible = False
'         Critnote.Show
'   End If
'ElseIf crdit.value = True Then
'   If vs1.Col = 1 Then
'         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
'         'MainMenu.Toolbar1.Visible = False
'         Creditnotefile.Show
'   End If
'ElseIf dbit.value = True Then
'   If vs1.Col = 1 Then
'         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
'         'MainMenu.Toolbar1.Visible = False
'         Debitnotefile.Show
'   End If
End If
End If
Screen.MousePointer = vbDefault
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If vs.Col = 4 Then
   SendKeys "{down}"
End If
End If
End Sub
Sub CalculateTotalDrCr()
On Error Resume Next
Dim Balance As Long
Dim dr1, cr1, prbal
Dim Str
Str = ""
dr1 = 0
cr1 = 0
txtClosing.Text = 0
txtcr.Text = 0
If RS.State = 1 Then RS.close
RS.Open "select Op,drcr from SLEDGER where " & stringyear & " and AgentName='" & cboParty.Text & "'", con
If RS.EOF = False Then
txtOp.Text = Format(RS.Fields(0).value, "0.00")
If UCase(RS.Fields("drcr").value) = UCase("dr") Then
cboop.Text = "Dr"
Else
cboop.Text = "Cr"
End If
Else
txtOp.Text = 0
End If
If cboop.Text = "Dr" Then
dr1 = (Val(txtOp.Text) + Val(vs1.TextMatrix(1, 4)))
cr1 = Val(vs1.TextMatrix(1, 5))
Else
cr1 = (Val(txtOp.Text) + Val(vs1.TextMatrix(1, 5)))
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
txtClosing.Text = (Val(txtClosing.Text) + Val(vs1.TextMatrix(I, 4)))
txtcr.Text = (Val(txtcr.Text) + Val(vs1.TextMatrix(I, 5)))
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
'---------------------------
End If
Next
txtClosing.Text = Format(txtClosing.Text, "0.00")
txtcr.Text = Format(txtcr.Text, "0.00")
If cboop.Text = "Dr" Then
txtClosing.Text = Format((CDbl(txtClosing.Text) + CDbl(txtOp.Text)), "0.00")
Else
txtcr.Text = Format((CDbl(txtcr.Text) + CDbl(txtOp.Text)), "0.00")
End If
txtBalance.Text = (Val(txtClosing.Text) - Val(txtcr.Text))
If Val(txtBalance.Text) < 1 Then
txtBalance.Text = (-1 * Val(txtBalance.Text))
closingcr.Text = "Cr"
Else
closingcr.Text = "Dr"
End If
txtBalance.Text = Format(txtBalance.Text, "0.00")

End Sub
Sub SaveDatainTempledger()

Dim d1 As Date



con.Execute "delete  from templedger1 where " & stringyear & ""
For I = 1 To vs1.rows - 1
If vs1.TextMatrix(I, 1) <> "" Then
con.Execute "INSERT INTO  templedger1(Party,dates,Billtype,Bill,Des,Dr,Cr,Balance,drcr,Party1,setupid,fyear)  values('" & Trim(cboParty) & "','" & Format(vs1.TextMatrix(I, 2), "MM/dd/yyyy") & "','" & vs1.TextMatrix(I, 0) & "', " & vs1.TextMatrix(I, 1) & ",'" & vs1.TextMatrix(I, 3) & "' ," & vs1.TextMatrix(I, 4) & "," & vs1.TextMatrix(I, 5) & "," & Val(vs1.TextMatrix(I, 6)) & ",'" & vs1.TextMatrix(I, 7) & "','" & vs1.TextMatrix(I, 8) & "'," & setupid & ",'" & session & "')"
End If
Next

Dim ff As New ADODB.Recordset
If ff.State = 1 Then ff.close
ff.Open "select Billtype,bill,dates,des,dr,cr,Balance,drcr,Party1 from templedger1 where " & stringyear & "  order by dates,bill", con
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
     vs1.TextMatrix(J, 8) = Format(ff.Fields(8).value, "0.00")
     ff.MoveNext
 End If
Next

End Sub
Private Sub cboParty_GotFocus()

Dim ph_rs As New ADODB.Recordset
cboParty.BackColor = &HFFFFC0


I = 1
If PopUpValue1 <> "" Then
cboParty.Text = PopUpValue1
End If




End Sub

Private Sub cboParty_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dr, cr As Double

If KeyCode = 27 Then Unload Me


If KeyCode = 113 Then
'-------------------------------
    din_ = False
    value = "select  rep as Reprasentative,Add1 As Address from rep order by rep"
    popuplistModel10 value, CON_blue
    set_focus = True
End If

If KeyCode = 115 Then
'-------------------------------
    din_ = True
    value = "SELECT Shipto,Shipto_City As City,Shipto_district as District,Shipto_States as States FROM INVOICEA_sp where len(Shipto)>0"
    popuplistModel10 value, con
    set_focus = True
End If



If KeyCode = 13 Then
If cboParty.Text = "" Then
  cboParty.SetFocus
  Exit Sub
End If

searchAllotment (cboParty.Text)
dataSearchingrid
cmdPrint.Enabled = True



dr = 0
cr = 0

For I = 1 To vs1.rows - 1
  dr = dr + Val(vs1.TextMatrix(I, 4))
  cr = cr + Val(vs1.TextMatrix(I, 5))
Next

drLebel.Caption = Format(dr, "0.00")
CrLebel.Caption = Format(cr, "0.00")
    
'txtdes.SetFocus
    
End If


If KeyCode = 116 Then
vs1.SetFocus
For J = 1 To vs1.rows - 1
   SendKeys "{down}"
   vs1.Row = J
Next
End If


End Sub
Sub dataSearchingrid()

Screen.MousePointer = vbHourglass
I = 1


If PopUpValue1 <> "" Then
vs1.Clear
vs1.rows = 1
fillGrid
End If

If cboParty.Text <> "" Then
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
    Set Del = con.Execute("delete from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.Text & "")
End Sub
Private Sub cmdDel_Click()
  
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
If Val(txtQty.Text) > 0 And txtdes.Text <> "" And cboParty.Text <> "" Then
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
    
    
    vs1.FormatString = "^Bill Type|^Bill|^Dates|Description|>Dr|>Cr|Balance|Dr/Cr|SchoolName"
    vs1.ColWidth(0) = 500
    vs1.ColWidth(1) = 1000
    vs1.ColWidth(2) = 1200
    vs1.ColWidth(3) = 2400
    vs1.ColWidth(4) = 1200
    vs1.ColWidth(5) = 1200
    vs1.ColWidth(6) = 1300
    vs1.ColWidth(7) = 500
    vs1.ColWidth(8) = 3000
    
    
   DoEvents

End Sub
Private Sub cmdModify_Click()


'''''''''''On Error GoTo aa1
''''''''''If MsgBox("Do U Want To Update ?", vbQuestion + vbYesNo) = vbYes Then
'''''''''''DelFunction
''''''''''CON.Execute "update ReceiveIssueParty set Dr=0,cr=0 where " & stringyear & " and RecNo=" & txtRecno.Text & ""
''''''''''
'''''''''''------------------------
''''''''''Set RS = New ADODB.Recordset
''''''''''RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.Text & "", CON, adOpenDynamic, adLockOptimistic
''''''''''If RS.EOF = False Then
'''''''''''maxId
'''''''''''RS.AddNew
''''''''''RS.Fields("RecNo").value = txtRecno.Text
''''''''''RS.Fields("Dates").value = RecDates.value
''''''''''RS.Fields("PartyName").value = cboParty.Text
''''''''''RS.Fields("Particullar").value = txtdes.Text
''''''''''If Receive.value = True Then
''''''''''RS.Fields("Dr").value = Val(txtQty.Text)
''''''''''Else
''''''''''RS.Fields("Cr").value = Val(txtQty.Text)
''''''''''End If
''''''''''RS.update
''''''''''End If
'''''''''''------------------------
''''''''''
'''''''''''SaveMain
''''''''''
''''''''''
''''''''''fillGrid
''''''''''CalculateTotalDrCr
''''''''''setwidth
''''''''''Call cmdRefresh_Click
''''''''''vs1.SetFocus
''''''''''For I = 1 To vs1.Rows - 1
''''''''''SendKeys "{down}"
''''''''''Next
''''''''''
''''''''''cmdModify.Enabled = False
''''''''''cmdDel.Enabled = False
''''''''''End If
'Exit Sub
'aa1:
'MsgBox "Record not Save !!", vbCritical
End Sub
Private Sub cmdRefresh_Click()
 
 
 Dim o As Object
 txtQty.Text = ""
 set_focus = False
 maxId
 cmdModify.Enabled = False
 cmdDel.Enabled = False
 cmdSave.Enabled = True
 
 
 Screen.MousePointer = vbDefault
 bb2 = False

End Sub
Private Sub cmdSave_Click()

'''''''''On Error GoTo aa:
'''''''''
'''''''''
'''''''''
'''''''''If cboParty.Text = "" Then
'''''''''MsgBox "Please Select Party Name !!", vbInformation
'''''''''Exit Sub
'''''''''End If
'''''''''
'''''''''If txtQty.Text = "" Then
'''''''''MsgBox "Please Enter Amount!!", vbInformation
'''''''''txtQty.SetFocus
'''''''''Exit Sub
'''''''''End If
'''''''''
'''''''''
'''''''''If RS.State = 1 Then RS.close
'''''''''RS.Open "select * from pass where pass='" & cp & "'", CON
'''''''''If RS.EOF = True Then
'''''''''MsgBox "Enter Valid Password !!", vbInformation
'''''''''Exit Sub
'''''''''End If
'''''''''
'''''''''If MsgBox("Do U Want To Save ?", vbInformation + vbYesNo) = vbYes Then
'''''''''aa1:
'''''''''SaveMain
'''''''''
'''''''''cboParty.SetFocus
'''''''''
'''''''''Call cmdRefresh_Click
'''''''''fillGrid
'''''''''
'''''''''cmdModify.Enabled = False
'''''''''cmdDel.Enabled = False
''''''''''----------------
'''''''''dataSearchingrid
''''''''''---------------
'''''''''
'''''''''
'''''''''End If
'''''''''Exit Sub
'''''''''aa:
'''''''''maxId
'''''''''GoTo aa1

End Sub
Sub SaveMain()
   
'''''''   maxId
'''''''    Set RS = New ADODB.Recordset
'''''''    RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.Text & "", CON, adOpenDynamic, adLockOptimistic
'''''''    If RS.EOF = True Then
'''''''       maxId
'''''''       RS.AddNew
'''''''       RS.Fields("RecNo").value = txtRecno.Text
'''''''       RS.Fields("Dates").value = RecDates.value
'''''''       RS.Fields("PartyName").value = cboParty.Text
'''''''       RS.Fields("Particullar").value = txtdes.Text
'''''''       If Receive.value = True Then
'''''''          RS.Fields("Dr").value = Val(txtQty.Text)
'''''''        Else
'''''''          RS.Fields("Cr").value = Val(txtQty.Text)
'''''''       End If
'''''''
'''''''    RS.update
'''''''    End If
End Sub
'''Sub search()
''''' If set_focus = True Then Exit Sub
''''' On Error Resume Next
'''''
'''''
'''''
'''''    If rss.State = 1 Then rss.close
'''''    rss.Open "select * from sledger where " & stringyear & " and AgentName=" & txtParty.Text & "", CON, adOpenDynamic, adLockOptimistic
'''''    If rss.EOF = 1 Then
'''''       txtRem.Text = RS.Fields("PartyRemarks").value & ""
'''''    End If
'''''
'''''
'''''
'''''
''''' If vs1.TextMatrix(vs1.RowSel, 0) = "J" Then
'''''    If RS.State = 1 Then RS.close
'''''    RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & vs1.TextMatrix(vs1.RowSel, 1) & "", CON, adOpenDynamic, adLockOptimistic
'''''    If RS.EOF = False Then
'''''       txtRecno.Text = RS.Fields("RecNo").value
'''''       RecDates.value = RS.Fields("Dates").value
'''''       cboParty.Text = RS.Fields("PartyName").value
'''''       txtdes.Text = RS.Fields("Particullar").value
'''''
'''''
'''''       If RS.Fields("Dr").value > 0 Then
'''''          Receive.value = True
'''''          txtQty.Text = RS.Fields("Dr").value
'''''        Else
'''''          Issue.value = True
'''''          txtQty.Text = RS.Fields("Cr").value
'''''       End If
'''''      End If
'''''   cmdSave.Enabled = False
'''''   cmdModify.Enabled = True
'''''   cmdDel.Enabled = True
'''''  Else
'''''   cmdModify.Enabled = False
'''''   cmdDel.Enabled = False
'''''   cmdSave.Enabled = True
'''''   txtdes.Text = ""
'''''   txtQty.Text = ""
'''''  End If
''End Sub
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
    SendKeys "{down}"
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
        'If RS.State = 1 Then RS.close
        'RS.Open "select * from pass where pass='" & cp & "'", CON
        'If RS.EOF = True Then
        '  If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
        '   Unload Me
        '   End If
        'Exit Sub
        'End If
        
        If Val(txtQty.Text) > 0 And txtdes.Text <> "" And cboParty.Text <> "" Then
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
    
   
    '==============
    SearchFa
    '==============
    setWidth
End Sub
Sub maxId()
  Dim rr As New ADODB.Recordset
  Set rr = New ADODB.Recordset
  rr.Open "select max(RecNo) from ReceiveIssueParty where " & stringyear & " ", con
  If IsNull(rr.Fields(0).value) Then
     txtRecno.Text = 1
     Else
     txtRecno.Text = rr.Fields(0).value + 1
  End If
End Sub

Private Sub Todate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    from_date = FromDate.value
    to_date = toDate.value
    fillGrid
    Frame1.Visible = False
 End If
End Sub

Private Sub txtQty_GotFocus()
   txtQty.BackColor = &HFFFFC0
End Sub

Private Sub txtQty_LostFocus()
  txtQty.BackColor = &HFFFFFF
End Sub
Private Sub txtRecno_KeyPress(KeyAscii As Integer)
   On Error Resume Next
  

  
  If KeyAscii = 13 Then
  
     If RS.State = 1 Then RS.close
     RS.Open "select * from receiveissueparty where " & stringyear & " and recno=" & txtRecno.Text & "", con
     If RS.EOF = False Then
      cboParty.Text = RS!partyname
      PopUpValue3 = cboParty.Text
      
      RecDates.value = RS.Fields("Dates").value
      txtdes.Text = RS.Fields("Particullar").value
      'txtRem.Text = RS.Fields("Remarks").Value
      If RS.Fields("Dr").value > 0 Then
          Receive.value = True
          txtQty.Text = RS.Fields("Dr").value
      Else
          Issue.value = True
          txtQty.Text = RS.Fields("Cr").value
      End If
      dataSearchingrid
     Else
       vs1.Clear
       setWidth
       txtQty.Text = ""
       txtdes.Text = ""
       cboParty.Text = ""
       txtOp.Text = ""
       txtBalance.Text = ""
     End If
  End If
End Sub
Private Sub txtSlipNo_GotFocus()
 txtSlipNo.BackColor = &HFFFFC0
End Sub
Private Sub txtSlipNo_LostFocus()
txtSlipNo.BackColor = &HFFFFFF
End Sub
Private Sub vs1_Click()
' search
End Sub
Private Sub vs1_DblClick()
set_focus = False
End Sub

Private Sub vs1_SelChange()
 'search
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
    cboParty.Text = vsop.TextMatrix(vsop.RowSel, 1)
    PopUpValue2 = cboParty.Text
    vs1.Clear
    fillGrid
    SaveDatainTempledger
    CalculateTotalDrCr
    setWidth
    PopUpValue1 = ""
    Opening.Tab = 1
End If
End Sub

