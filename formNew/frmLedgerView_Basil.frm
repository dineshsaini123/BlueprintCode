VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLedgerView_Basil 
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   10185
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
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2805
      Visible         =   0   'False
      Width           =   870
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
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3225
      Visible         =   0   'False
      Width           =   870
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
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3645
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
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1845
      Visible         =   0   'False
      Width           =   840
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
      Left            =   12810
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4125
      Visible         =   0   'False
      Width           =   840
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
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1845
      Visible         =   0   'False
      Width           =   810
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
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2325
      Visible         =   0   'False
      Width           =   825
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
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4605
      Visible         =   0   'False
      Width           =   735
   End
   Begin TabDlg.SSTab Opening 
      Height          =   4185
      Left            =   12150
      TabIndex        =   0
      Top             =   1605
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
      TabPicture(0)   =   "frmLedgerView_Basil.frx":0000
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
      TabPicture(1)   =   "frmLedgerView_Basil.frx":001C
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
      TabPicture(2)   =   "frmLedgerView_Basil.frx":0038
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
         TabIndex        =   19
         Top             =   7755
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Search FA"
         Height          =   210
         Left            =   4050
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   7755
         Visible         =   0   'False
         Width           =   1395
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
         TabIndex        =   16
         Top             =   660
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox selectAll 
         Caption         =   "Select All"
         Height          =   300
         Left            =   -65880
         TabIndex        =   15
         Top             =   2295
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.ListBox cboPartyList 
         Appearance      =   0  'Flat
         Height          =   2730
         Left            =   -67920
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   2595
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.CommandButton cmddewali 
         Caption         =   "Refresh Diwali Head"
         Height          =   255
         Left            =   -66645
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1875
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
         TabIndex        =   12
         Top             =   2295
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox cboStation 
         Height          =   315
         Left            =   -67905
         TabIndex        =   11
         Top             =   2280
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.ComboBox closingcr 
         Height          =   315
         Left            =   -65265
         Style           =   1  'Simple Combo
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1560
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton cmdShow1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -70260
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.ComboBox cboop 
         Height          =   315
         Left            =   -65265
         Style           =   1  'Simple Combo
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1365
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
         TabIndex        =   5
         Top             =   630
         Visible         =   0   'False
         Width           =   1335
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   1830
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   675
      End
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
         TabIndex        =   1
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
         TabIndex        =   20
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   1695
         Width           =   6975
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
         TabIndex        =   33
         Top             =   6960
         Width           =   1275
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame panel 
      Caption         =   "Ledger View"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   7455
      Left            =   120
      TabIndex        =   47
      Top             =   180
      Width           =   9795
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         Height          =   660
         Left            =   6480
         Picture         =   "frmLedgerView_Basil.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Width           =   1005
      End
      Begin VB.ComboBox cboParty 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   360
         Left            =   195
         Style           =   1  'Simple Combo
         TabIndex        =   48
         Top             =   480
         Width           =   6015
      End
      Begin VSFlex7Ctl.VSFlexGrid vs1 
         Height          =   6315
         Left            =   180
         TabIndex        =   49
         Top             =   960
         Width           =   9570
         _cx             =   16880
         _cy             =   11139
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
   Begin VB.Frame f1 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   120
      TabIndex        =   44
      Top             =   540
      Visible         =   0   'False
      Width           =   90
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
         TabIndex        =   46
         Top             =   60
         Width           =   3375
      End
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
         TabIndex        =   45
         Top             =   45
         Width           =   3360
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   15
         Stretch         =   -1  'True
         Top             =   75
         Width           =   10155
      End
   End
End
Attribute VB_Name = "frmLedgerView_Basil"
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
RS.Open "select distinct(SUBLEDGER) from SLEDGER where " & stringyear & " and DISTCODE='" & cboStation.Text & "'", con
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
    f.Open "select AMOUNT,text,INVOICENO from CASHc_basil where " & stringyear & " and TEXT = '" & "DIWALI SPECIAL" & "' and AMOUNT>0", con
    While f.EOF = False
        con.Execute "update CASHA_basil set t2='" & f.Fields("amount").value & "' where " & stringyear & " and INVOICENO=" & f.Fields("INVOICENO").value & ""
        f.MoveNext
    Wend
    If f.State = 1 Then f.close
    f.Open "select AMOUNT,text,INVOICENO from CASHC_basil where " & stringyear & " and TEXT = '" & "DIWALI SPECIAL" & "' and AMOUNT>0", con
    While f.EOF = False
        con.Execute "update CASHA_basil set t2='" & f.Fields("amount").value & "' where " & stringyear & " and INVOICENO=" & f.Fields("INVOICENO").value & ""
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
con.Execute "delete from templedger1 where " & stringyear

If RS.State = 1 Then RS.close
RS.Open "select subledger from SLEDGER where " & stringyear & " and subledger = '" + Trim(cboParty.Text) + "'", con

While RS.EOF = False

'==Code For Opening=============================================
con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype,fyear,setupid)  SELECT op,drcr,subledger,'Opening',fyear,setupid from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'   group by op,subledger,drcr HAVING  op <> 0;"
If rs1.State = 1 Then rs1.close
rs1.Open "SELECT op,drcr from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'", con
If Not IsNull(rs1.Fields(0).value) Then
   op = Val(rs1.Fields(0).value)
   drcr = rs1.Fields(1).value
Else
   op = 0
End If

'==============================================
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER,fyear,setupid from CASHA_basil where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',BAA,netamount,SUBLEDGER,fyear,setupid from CREDITA where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER,fyear,setupid from CASHA_basil where  " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD,fyear,setupid from dnfa where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party,fyear,setupid) Select cnd,'CN',cnn,'Credit Note',na,'0',psld,fyear,setupid from Cnf1a where " & stringyear & " and  psld='" & RS.Fields(0).value & "'"
con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid) Select dates,'J',Recno,Particullar,Dr,CR,PartyName,fyear,setupid from ReceiveIssueParty where " & stringyear & " and PartyName='" & RS.Fields(0).value & "' order by dates,recno"
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
   
   If MsgBox("Want to Set ?", vbQuestion + vbYesNo) = vbYes Then
        
   Screen.MousePointer = vbHourglass
   'cmdShow1.Visible = True
         
         
   If sales.value = True Then
        
        For J = 1 To vs.Rows - 1
          If vs.TextMatrix(J, 5) = True Then
             con.Execute "update CASHA_basil set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
             con.Execute "update CASHA_basil set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
   
  ElseIf credit.value = True Then
  
        For J = 1 To vs.Rows - 1
          If vs.TextMatrix(J, 5) = True Then
            con.Execute "update CREDITA set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
            con.Execute "update CREDITA set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
  
  
  ElseIf cash.value = True Then
        
        For J = 1 To vs.Rows - 1
          If vs.TextMatrix(J, 5) = True Then
            con.Execute "update CASHA_basil set BAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
            con.Execute "update CASHA_basil set BAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
  
  ElseIf crdit.value = True Then
  
        For J = 1 To vs.Rows - 1
          If vs.TextMatrix(J, 5) = True Then
            con.Execute "update cnf1a set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and cnn=" & vs.TextMatrix(J, 1) & ""
            Else
            con.Execute "update cnf1a set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and cnn=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
  
  
  ElseIf dbit.value = True Then
  
        For J = 1 To vs.Rows - 1
          If vs.TextMatrix(J, 5) = True Then
            con.Execute "update dnfa set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and dnn=" & vs.TextMatrix(J, 1) & ""
            Else
            con.Execute "update dnfa set bAuthorized=" & vs.TextMatrix(J, 5) & " where " & stringyear & " and dnn=" & vs.TextMatrix(J, 1) & ""
          End If
        Next
   
   
  End If
   
   
   End If
   
   
 Screen.MousePointer = vbDefault
End Sub


Sub SearchFa()

'    '================
     If RS.State = 1 Then RS.close
     RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,baa from casha_basilRet where " & stringyear & " and SUBLEDGER='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
     If RS.EOF = False Then
        vs1.Rows = vs1.Rows + RS.RecordCount
        For I = I To vs1.Rows - 1
         
        If RS.EOF = False Then
         vs1.TextMatrix(I, 0) = "Ret"
         vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
         vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
         vs1.TextMatrix(I, 3) = "Return"
         vs1.TextMatrix(I, 4) = Format(RS.Fields("baa").value, "0.00")
         vs1.TextMatrix(I, 5) = Format(RS.Fields("netamount").value, "0.00")
         RS.MoveNext
       End If
    Next
    End If
    If RS.State = 1 Then RS.close
    RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,netamount,baa,t2 from CASHA_basil where  " & stringyear & " and SUBLEDGER='" & cboParty.Text & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)", con
    If RS.EOF = False Then
     vs1.Rows = vs1.Rows + RS.RecordCount
     For I = I To vs1.Rows - 1
    If RS.EOF = False Then
      vs1.TextMatrix(I, 0) = "Est"
      vs1.TextMatrix(I, 1) = RS.Fields("INVOICENO").value
      vs1.TextMatrix(I, 2) = RS.Fields("INVOICEDATE").value
      If IsNull(RS.Fields("t2").value) Then
         vs1.TextMatrix(I, 3) = "Estimate"
      Else
         vs1.TextMatrix(I, 3) = "Cash Memo" & " " & RS.Fields("t2").value & " DS"
         
      End If
      vs1.TextMatrix(I, 4) = Format(RS.Fields("netamount").value, "0.00")
      vs1.TextMatrix(I, 5) = Format(RS.Fields("baa").value, "0.00")
      RS.MoveNext
    End If
    Next
    End If
'===================
    If RS.State = 1 Then RS.close
    RS.Open "select cnn,cnd,na from Cnf1a where  psld='" & cboParty.Text & "'", con
    If RS.EOF = False Then
     vs1.Rows = vs1.Rows + RS.RecordCount
     For I = I To vs1.Rows - 1
      
    If RS.EOF = False Then
    
      vs1.TextMatrix(I, 0) = "CN"
      vs1.TextMatrix(I, 1) = RS.Fields("cnn").value
      vs1.TextMatrix(I, 2) = RS.Fields("cnd").value
      vs1.TextMatrix(I, 3) = "Credit Note"
      vs1.TextMatrix(I, 5) = Format(RS.Fields("na").value, "0.00")
      vs1.TextMatrix(I, 4) = 0
      RS.MoveNext
    
    End If
    
    Next
    End If
     '===================
    If RS.State = 1 Then RS.close
    RS.Open "select dnn,dnd,psld,na from dnfa where  " & stringyear & " and psld='" & cboParty.Text & "'", con
    If RS.EOF = False Then
     vs1.Rows = vs1.Rows + RS.RecordCount
     For I = I To vs1.Rows - 1
    If RS.EOF = False Then
      vs1.TextMatrix(I, 0) = "DN"
      vs1.TextMatrix(I, 1) = RS.Fields("dnn").value
      vs1.TextMatrix(I, 2) = RS.Fields("dnd").value
      vs1.TextMatrix(I, 3) = "Debit Note"
      vs1.TextMatrix(I, 4) = Format(RS.Fields("na").value, "0.00")
      vs1.TextMatrix(I, 5) = 0
      RS.MoveNext
    End If
    Next
    End If
    vs1.FormatString = "^Bill Type|^Bill|^Date|<Description|>Dr|>Cr"
    setwidth
End Sub
Sub CityWiseStatement()
       Dim op, drcr
       Dim s As String
       s = ""
       Dim rs1 As New ADODB.Recordset
       con.Execute "delete from templedger1"
       If RS.State = 1 Then RS.close
       If cboStation.Text <> "" And txtalfa.Text = "" Then
       For I = 0 To cboPartyList.ListCount - 1
        If cboPartyList.Selected(I) = True Then
        If s = "" Then
          s = "SUBLEDGER " & " = " & "'" & cboPartyList.List(I) & "'"
        Else
          s = s & " or " & "SUBLEDGER " & " = " & "'" & cboPartyList.List(I) & "'"
        End If
        End If
       Next
       
       If s = "" Then
        If RS.State = 1 Then RS.close
        RS.Open "select subledger from SLEDGER where " & stringyear & " and DISTCODE = '" & cboStation.Text & "'", con
       Else
        If RS.State = 1 Then RS.close
        RS.Open "select subledger from SLEDGER where " & stringyear & " and " & s, con
       End If
       
       ElseIf txtalfa.Text <> "" And cboStation.Text = "" Then
        RS.Open "select subledger from SLEDGER where " & stringyear & " and Subledger like '" + Trim(txtalfa.Text) + "%'", con
       Else
         Exit Sub
       End If
       While RS.EOF = False
           '==Code For Opening=============================================
            con.Execute "INSERT INTO templedger1 (Balance,drcr,party,billtype,fyear,setupid)  SELECT op,drcr,subledger,'Opening',fyear,setupid  from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'   group by op,subledger,drcr HAVING  op <> 0;"
            If rs1.State = 1 Then rs1.close
            rs1.Open "SELECT op,drcr from sledger where " & stringyear & " and subledger = '" + RS.Fields(0).value + "'", con
            If Not IsNull(rs1.Fields(0).value) Then
               op = Val(rs1.Fields(0).value)
               drcr = rs1.Fields(1).value
            Else
               op = 0
            End If
           '==============================================
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales',netamount,BAA,SUBLEDGER,fyear,setupid from CASHA_basil where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',baa,netamount,SUBLEDGER,fyear,setupid  from CREDITA where " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid) Select INVOICEDATE,'C/M',INVOICENO,'Cash Memo',netamount,baa,SUBLEDGER,fyear,setupid from CASHA_basil where  " & stringyear & " and SUBLEDGER='" & RS.Fields(0).value & "' and ((netamount-BAA)>0 or (netamount-BAA)<0)"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid) select dnd,'DN',dnn,'Debit Note',na,'0',PSLD from dnfa,fyear,setupid where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,cr,dr,Party,fyear,setupid) Select cnd,'CN',cnn,'Credit Note',na,'0',psld,fyear,setupid  from Cnf1a where  " & stringyear & " and psld='" & RS.Fields(0).value & "'"
            con.Execute "INSERT INTO templedger1 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid) Select dates,'J',Recno,Particullar,Dr,CR,PartyName,fyear,setupid from ReceiveIssueParty where " & stringyear & " and PartyName='" & RS.Fields(0).value & "' order by dates,recno"
          '===============================================================
          If op <> 0 Then
           con.Execute "Update templedger1 set Balance=" & op & ",drcr= '" & drcr & "' where " & stringyear & " and party = '" + RS.Fields(0).value + "' and Billtype<>'Opening'"
          End If
          '===============================================================
           RS.MoveNext
       Wend
       DoEvents
       
       DSNNew
       
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
   RS.Open "select subledger from sledger", con
   While RS.EOF = False
       
       aa = InStr(RS(0), " ")
       partyname = Mid(RS(0), aa)
       pcode = Mid(RS(0), 1, aa)
       
       con.Execute "update  Sledger  set party='" & Trim(partyname) & "',code='" & Trim(pcode) & "' where " & stringyear & " and subledger='" & RS(0) & "'"
       
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
   
   For I = 1 To vsop.Rows - 1
       If vsop.TextMatrix(I, 1) <> "" Then
          con.Execute "update SLEDGER set op=" & CDbl(vsop.TextMatrix(I, 2)) & ",drcr='" & vsop.TextMatrix(I, 3) & "' where " & stringyear & " and SUBLEDGER='" & vsop.TextMatrix(I, 1) & "'"
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
Dim total As String
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
 
Me.Left = 100
Me.Top = 100
Me.Width = 10200
Me.Height = 8200
 
' Me.WindowState = 2
   
   
Dim rs_1 As New ADODB.Recordset
If rs_1.State = 1 Then rs_1.close
rs_1.Open "select * from pass where pass='" & strledger & "'", con
If rs_1.EOF = True Then
   'txtRem.Visible = False
   cmdShow1.Visible = False
   'txtRem.Visible = True
   txtRem.Enabled = False
Else
   'txtRem.Visible = True
   txtRem.Enabled = True
   'cmdShow1.Visible = True
End If



   
  
End Sub

Private Sub Form_Load()

cboParty.Text = ""
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""



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
setwidth

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
RS.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdTable
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

        Dim fillvs As New ADODB.Recordset
        If fillvs.State = 1 Then fillvs.close
        'fillvs.Open "select DISTCODE as City,SUBLEDGER as Party,op,drcr from closing where gledger='SUNDRY DEBTORS'", con
        fillvs.Open "SELECT SLEDGER.DISTCODE,SLEDGER.SUBLEDGER,SLEDGER.OP,SLEDGER.drcr,(Sum(templedger1.Dr)-Sum(templedger1.Cr)) AS bal1 FROM SLEDGER LEFT JOIN templedger1 ON SLEDGER.SUBLEDGER = templedger1.Party where  " & stringyear & " and gledger='SUNDRY DEBTORS' GROUP BY SLEDGER.SUBLEDGER,SLEDGER.DISTCODE,[SLEDGER.OP], SLEDGER.drcr, SLEDGER.gledger", con

        If fillvs.EOF = False Then
            vsop.Rows = fillvs.RecordCount
            For I = 1 To vsop.Rows - 1
              vsop.TextMatrix(I, 0) = fillvs(0) & ""
              vsop.TextMatrix(I, 1) = fillvs(1)
              vsop.TextMatrix(I, 2) = Format(fillvs(2), "0.00")
              vsop.TextMatrix(I, 3) = fillvs(3) & ""

              If Not IsNull(fillvs(4)) Then

                     If vsop.TextMatrix(I, 3) = "Cr" Then
                         vsop.TextMatrix(I, 4) = ((-1 * (vsop.TextMatrix(I, 2))) + fillvs(4))
                         If vsop.TextMatrix(I, 4) < 0 Then
                            vsop.TextMatrix(I, 4) = Format((-1 * Val(vsop.TextMatrix(I, 4))), "0.00")
                            vsop.TextMatrix(I, 5) = "Cr"
                         Else
                            vsop.TextMatrix(I, 5) = "Dr"
                            vsop.TextMatrix(I, 4) = Format(Val(vsop.TextMatrix(I, 4)), "0.00")
                         End If

                     Else
                         vsop.TextMatrix(I, 4) = ((Val(vsop.TextMatrix(I, 2))) + fillvs(4))
                         If vsop.TextMatrix(I, 4) < 0 Then
                            vsop.TextMatrix(I, 4) = Format((-1 * Val(vsop.TextMatrix(I, 4))), "0.00")
                            vsop.TextMatrix(I, 5) = "Cr"
                         Else
                            vsop.TextMatrix(I, 5) = "Dr"
                            vsop.TextMatrix(I, 4) = Format(Val(vsop.TextMatrix(I, 4)), "0.00")
                         End If


                     End If
              End If


              fillvs.MoveNext
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
       value = "select SUBLEDGER from CASHA_basil  where " & stringyear & " order by SUBLEDGER"
       popuplist10 value, con
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
     con.Execute "update sledger set PartyRemarks = '" & txtRem.Text & "' where " & stringyear & " and subledger='" & cboParty.Text & "'"
     
  End If
  End If
End Sub
Private Sub Unautho_Click()
If Unautho.value = True Then
'    Call cmdShow_Click
End If
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
Screen.MousePointer = vbHourglass
If KeyCode = 13 Then
If sales.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         ''MainMenu.Toolbar1.Visible = False
         invoice.Show  '    sales
   End If
ElseIf cash.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         ''MainMenu.Toolbar1.Visible = False
         countersale.Show  '
   End If
ElseIf credit.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         ''MainMenu.Toolbar1.Visible = False
         Critnote.Show
   End If
ElseIf crdit.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         ''MainMenu.Toolbar1.Visible = False
         Creditnotefile.Show
   End If
ElseIf dbit.value = True Then
   If vs.Col = 1 Then
         inviceNo = vs.TextMatrix(vs.RowSel, 1)
         ''MainMenu.Toolbar1.Visible = False
         Debitnotefile.Show
   End If
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
RS.Open "select Op,drcr from SLEDGER where " & stringyear & " and SUBLEDGER='" & cboParty.Text & "'", con
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
For I = 1 To vs1.Rows - 1
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
con.Execute "delete from templedger1 where " & stringyear
For I = 1 To vs1.Rows - 1
If vs1.TextMatrix(I, 1) <> "" Then
con.Execute "INSERT INTO  templedger1(dates,Billtype,Bill,Des,Dr,Cr,Balance,drcr,fyear,setupid)  values('" & vs1.TextMatrix(I, 2) & "','" & vs1.TextMatrix(I, 0) & "', " & vs1.TextMatrix(I, 1) & ",'" & vs1.TextMatrix(I, 3) & "' ," & vs1.TextMatrix(I, 4) & "," & vs1.TextMatrix(I, 5) & "," & Val(vs1.TextMatrix(I, 6)) & ",'" & vs1.TextMatrix(I, 7) & "','" & session & "', '" & setupid & "')"
End If
Next
Dim ff As New ADODB.Recordset
If ff.State = 1 Then ff.close
ff.Open "select Billtype,bill,dates,des,dr,cr,Balance,drcr from templedger1 where " & stringyear & "  order by dates,bill", con
vs1.Rows = ff.RecordCount + 1
For J = 1 To vs1.Rows - 1
 If ff.EOF = False Then
     vs1.TextMatrix(J, 0) = ff.Fields(0).value
     vs1.TextMatrix(J, 1) = ff.Fields(1).value
     vs1.TextMatrix(J, 2) = ff.Fields(2).value
     vs1.TextMatrix(J, 3) = ff.Fields(3).value
     vs1.TextMatrix(J, 4) = Format(ff.Fields(4).value, "0.00")
     vs1.TextMatrix(J, 5) = Format(ff.Fields(5).value, "0.00")
     vs1.TextMatrix(J, 6) = Format(ff.Fields(6).value, "0.00")
     ff.MoveNext
 End If
Next
End Sub
Private Sub cboParty_GotFocus()

Dim ph_rs As New ADODB.Recordset
cboParty.BackColor = &HFFFFC0

If search_v = False Then

I = 1
If PopUpValue3 = "" Then
PopUpValue2 = cboParty.Text
End If

If PopUpValue3 <> "" Then
cboParty.Text = PopUpValue3

Set ph_rs = New ADODB.Recordset
ph_rs.Open "select phone,PartyRemarks,MOBILE from sledger where " & stringyear & " and subledger='" & cboParty.Text & "'", con, adOpenKeyset, adLockReadOnly
If ph_rs.EOF = False Then
   phone.Caption = ph_rs(0) & "," & ph_rs!mobile
   txtRem.Text = ph_rs.Fields("PartyRemarks").value & ""
Else
   phone.Caption = ""
   txtRem.Text = ""
End If
End If


'--------------------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------------------

I = 1
If PopUpValue2 = "" Then
PopUpValue2 = cboParty.Text
End If

If PopUpValue2 <> "" Then
cboParty.Text = PopUpValue2
Set ph_rs = New ADODB.Recordset
ph_rs.Open "select phone,PartyRemarks,MOBILE from sledger where " & stringyear & " and subledger='" & cboParty.Text & "'", con, adOpenKeyset, adLockReadOnly
If ph_rs.EOF = False Then
   phone.Caption = ph_rs(0) & "," & ph_rs!mobile
   txtRem.Text = ph_rs.Fields("PartyRemarks").value & ""
Else
   phone.Caption = ""
   txtRem.Text = ""
End If
End If




End If


End Sub

Private Sub cboParty_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dr, cr As Double

If KeyCode = 114 Then
  search_v = True
Else
  search_v = False
End If


If search_v = False Then

If KeyCode = 113 Then
    value = "select distinct(Party),Code,subledger from SLEDGER where " & stringyear & " and gledger='SUNDRY DEBTORS' order by party"
    popuplist_client value, CCON
    set_focus = True
End If

If KeyCode = 13 Then
If cboParty.Text = "" Then
  cboParty.SetFocus
  Exit Sub
End If

dataSearchingrid
cmdPrint.Enabled = True



dr = 0
cr = 0

For I = 1 To vs1.Rows - 1
  dr = dr + Val(vs1.TextMatrix(I, 4))
  cr = cr + Val(vs1.TextMatrix(I, 5))
Next

drLebel.Caption = Format(dr, "0.00")
CrLebel.Caption = Format(cr, "0.00")
    
'txtdes.SetFocus
    
End If

'========================================================================================
Else
'========================================================================================


If KeyCode = 114 Then
    value = "SELECT  trim(mid( SLEDGER.SUBLEDGER,instr(SLEDGER.SUBLEDGER,',')+1)) as city,SUBLEDGER,Code FROM SLEDGER " & _
    "where " & stringyear & " and  instr(SLEDGER.SUBLEDGER,',')>0 and gledger='SUNDRY DEBTORS' order by trim(mid( SLEDGER.SUBLEDGER,instr(SLEDGER.SUBLEDGER,',')+1))"
    popuplistModel10 value, con
    set_focus = True
End If

If KeyCode = 13 Then
If cboParty.Text = "" Then
  cboParty.SetFocus
  Exit Sub
End If




dataSearchingrid
cmdPrint.Enabled = True

dr = 0
cr = 0

For I = 1 To vs1.Rows - 1
  dr = dr + Val(vs1.TextMatrix(I, 4))
  cr = cr + Val(vs1.TextMatrix(I, 5))
Next

drLebel.Caption = Format(dr, "0.00")
CrLebel.Caption = Format(cr, "0.00")
    
txtdes.SetFocus

End If

End If



If KeyCode = 116 Then
vs1.SetFocus
For J = 1 To vs1.Rows - 1
   SendKeys "{down}"
   vs1.Row = J
Next
End If


If KeyCode = 27 Then
   Unload Me
End If
  


End Sub
Sub dataSearchingrid()
Screen.MousePointer = vbHourglass
I = 1


If PopUpValue3 <> "" Then
vs1.Clear
vs1.Rows = 1
fillGrid
End If
If cboParty.Text <> "" Then
SaveDatainTempledger
CalculateTotalDrCr
End If
setwidth
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
Sub setwidth()
    vs1.FormatString = "^Bill Type|^Bill|^Dates|Description|>Dr|>Cr|Balance|Dr/Cr"
    vs1.ColWidth(0) = 500
    vs1.ColWidth(1) = 1000
    vs1.ColWidth(2) = 1000
    vs1.ColWidth(3) = 2400
    vs1.ColWidth(4) = 1200
    vs1.ColWidth(5) = 1200
    vs1.ColWidth(6) = 1300
    vs1.ColWidth(7) = 500
    
   DoEvents

End Sub
Private Sub cmdModify_Click()

'On Error GoTo aa1
If MsgBox("Do U Want To Update ?", vbQuestion + vbYesNo) = vbYes Then
'DelFunction
con.Execute "update ReceiveIssueParty set Dr=0,cr=0 where " & stringyear & " and RecNo=" & txtRecno.Text & ""

'------------------------
Set RS = New ADODB.Recordset
RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.Text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
'maxId
'RS.AddNew
RS.Fields("RecNo").value = txtRecno.Text
RS.Fields("Dates").value = RecDates.value
RS.Fields("PartyName").value = cboParty.Text
RS.Fields("Particullar").value = txtdes.Text
If Receive.value = True Then
RS.Fields("Dr").value = Val(txtQty.Text)
Else
RS.Fields("Cr").value = Val(txtQty.Text)
End If
RS.update
End If
'------------------------

'SaveMain


fillGrid
CalculateTotalDrCr
setwidth
Call cmdRefresh_Click
vs1.SetFocus
For I = 1 To vs1.Rows - 1
SendKeys "{down}"
Next

cmdModify.Enabled = False
cmdDel.Enabled = False
End If
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
Private Sub CmdSave_Click()

On Error GoTo aa:



If cboParty.Text = "" Then
MsgBox "Please Select Party Name !!", vbInformation
Exit Sub
End If

If txtQty.Text = "" Then
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
    RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & txtRecno.Text & "", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
       maxId
       RS.AddNew
       RS.Fields("RecNo").value = txtRecno.Text
       RS.Fields("Dates").value = RecDates.value
       RS.Fields("PartyName").value = cboParty.Text
       RS.Fields("Particullar").value = txtdes.Text
       If Receive.value = True Then
          RS.Fields("Dr").value = Val(txtQty.Text)
        Else
          RS.Fields("Cr").value = Val(txtQty.Text)
       End If
    
       RS.Fields("fyear").value = session
       RS.Fields("setupid").value = setupid
    
    RS.update
    End If
End Sub
Sub search()
 If set_focus = True Then Exit Sub
 On Error Resume Next
 
 
 
    If rss.State = 1 Then rss.close
    rss.Open "select * from sledger where " & stringyear & " and subledger=" & txtParty.Text & "", con, adOpenDynamic, adLockOptimistic
    If rss.EOF = 1 Then
       txtRem.Text = RS.Fields("PartyRemarks").value & ""
    End If

 
 
 
 If vs1.TextMatrix(vs1.RowSel, 0) = "J" Then
    If RS.State = 1 Then RS.close
    RS.Open "select * from ReceiveIssueParty where " & stringyear & " and RecNo=" & vs1.TextMatrix(vs1.RowSel, 1) & "", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       txtRecno.Text = RS.Fields("RecNo").value
       RecDates.value = RS.Fields("Dates").value
       cboParty.Text = RS.Fields("PartyName").value
       txtdes.Text = RS.Fields("Particullar").value
       
       
       If RS.Fields("Dr").value > 0 Then
          Receive.value = True
          txtQty.Text = RS.Fields("Dr").value
        Else
          Issue.value = True
          txtQty.Text = RS.Fields("Cr").value
       End If
      End If
   cmdSave.Enabled = False
   cmdModify.Enabled = True
   cmdDel.Enabled = True
  Else
   cmdModify.Enabled = False
   cmdDel.Enabled = False
   cmdSave.Enabled = True
   txtdes.Text = ""
   txtQty.Text = ""
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
    For I = 1 To vs1.Rows - 1
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
        If RS.State = 1 Then RS.close
        RS.Open "select * from pass where pass='" & cp & "'", con
        If RS.EOF = True Then
          If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
           Unload Me
           End If
        Exit Sub
        End If
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
    setwidth
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
     RS.Open "select * from receiveissueparty where " & stringyear & " and recno=" & txtRecno.Text & "", con, adOpenKeyset, adLockReadOnly
     If RS.EOF = False Then
      cboParty.Text = RS!partyname
      PopUpValue3 = cboParty.Text
      
      RecDates.value = RS.Fields("Dates").value
      txtdes.Text = RS.Fields("Particullar").value
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
       setwidth
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
 search
End Sub
Private Sub vs1_DblClick()
set_focus = False
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
   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         invoice.Show
   'End If

ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "Ret" Then

   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         'frmCashCounterRet.Show
         frmBasilSales_Ret.Show
   'End If

ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "CN" Then

   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         Creditnotefile.Show
   'End If

ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "DN" Then

   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         Debitnotefile.Show
   'End If

ElseIf vs1.TextMatrix(vs1.RowSel, 0) = "Est" Then

   'If vs1.col = 1 Then
         inviceNo = vs1.TextMatrix(vs1.RowSel, 1)
         'MainMenu.Toolbar1.Visible = False
         frmBasilSales.Show
         'countersale.Show
   'End If

End If


End If


If KeyCode = 112 Then
   txtdes.SetFocus
End If

Screen.MousePointer = vbDefault

End Sub

Private Sub vs1_SelChange()
 search
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
    setwidth
    PopUpValue1 = ""
    Opening.Tab = 1
End If
End Sub


