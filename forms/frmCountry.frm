VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCountry 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5760
   ClientLeft      =   2775
   ClientTop       =   2415
   ClientWidth     =   9225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   405
      TabIndex        =   1
      Top             =   4590
      Width           =   7575
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6255
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1230
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4515
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7964
      _Version        =   393216
      Tabs            =   8
      Tab             =   3
      TabsPerRow      =   8
      TabHeight       =   520
      BackColor       =   8454016
      TabCaption(0)   =   "Country"
      TabPicture(0)   =   "frmCountry.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Container"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "State"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "District"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "City/Town"
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "4"
      TabPicture(4)   =   "frmCountry.frx":001C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "5"
      TabPicture(5)   =   "frmCountry.frx":0038
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "6"
      TabPicture(6)   =   "frmCountry.frx":0054
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame6"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "7"
      TabPicture(7)   =   "frmCountry.frx":0070
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame7"
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   4155
         Left            =   -74955
         TabIndex        =   52
         Top             =   315
         Visible         =   0   'False
         Width           =   9045
         Begin VB.TextBox txtPan 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2475
            MaxLength       =   30
            TabIndex        =   60
            Top             =   3600
            Width           =   2625
         End
         Begin VB.TextBox txtAddress1 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2475
            MaxLength       =   40
            TabIndex        =   55
            Top             =   1530
            Width           =   3570
         End
         Begin VB.TextBox txtAddress2 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2475
            MaxLength       =   40
            TabIndex        =   56
            Top             =   1890
            Width           =   3570
         End
         Begin VB.TextBox txtCity 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2475
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   2250
            Width           =   3540
         End
         Begin VB.TextBox txtCityID1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   6030
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   68
            Top             =   2250
            Width           =   1110
         End
         Begin VB.TextBox txtpin 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   4680
            MaxLength       =   30
            TabIndex        =   58
            Top             =   2610
            Width           =   1320
         End
         Begin VB.TextBox txtPhone1 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2475
            MaxLength       =   30
            TabIndex        =   59
            Top             =   3285
            Width           =   3525
         End
         Begin VB.TextBox txtDist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2475
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   2610
            Width           =   1785
         End
         Begin VB.TextBox txtState 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2475
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   2970
            Width           =   3540
         End
         Begin VB.TextBox txtAuther 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2475
            MaxLength       =   30
            TabIndex        =   54
            Top             =   1155
            Width           =   2580
         End
         Begin VB.TextBox txtAuthId 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2475
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   810
            Width           =   2595
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pan :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   25
            Left            =   1485
            TabIndex        =   75
            Top             =   3645
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   24
            Left            =   1485
            TabIndex        =   74
            Top             =   1530
            Width           =   960
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "City/Town :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   23
            Left            =   1485
            TabIndex        =   73
            Top             =   2250
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "District :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   22
            Left            =   1485
            TabIndex        =   72
            Top             =   2610
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sate :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   21
            Left            =   1485
            TabIndex        =   71
            Top             =   2970
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pin :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   20
            Left            =   4275
            TabIndex        =   70
            Top             =   2655
            Width           =   510
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   19
            Left            =   1485
            TabIndex        =   69
            Top             =   3285
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Auther :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   285
            Index           =   18
            Left            =   1485
            TabIndex        =   64
            Top             =   1170
            Width           =   1275
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Auther Id  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   270
            Index           =   17
            Left            =   1485
            TabIndex        =   63
            Top             =   810
            Width           =   1290
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3975
         Left            =   -74955
         TabIndex        =   47
         Top             =   315
         Visible         =   0   'False
         Width           =   9000
         Begin VB.TextBox txtBookId 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1530
            Width           =   2595
         End
         Begin VB.TextBox txtBookType 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   48
            Top             =   1875
            Width           =   2580
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "BookType-Id :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   16
            Left            =   1260
            TabIndex        =   51
            Top             =   1530
            Width           =   1245
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Book Type :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Index           =   15
            Left            =   1260
            TabIndex        =   50
            Top             =   1890
            Width           =   1275
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3975
         Left            =   -74955
         TabIndex        =   42
         Top             =   315
         Visible         =   0   'False
         Width           =   9000
         Begin VB.TextBox txtDepartment 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   44
            Top             =   1875
            Width           =   2580
         End
         Begin VB.TextBox txtDepartmentId 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1530
            Width           =   2595
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Department :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Index           =   14
            Left            =   1260
            TabIndex        =   46
            Top             =   1905
            Width           =   1275
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Department-Id  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   13
            Left            =   1260
            TabIndex        =   45
            Top             =   1530
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3975
         Left            =   -74955
         TabIndex        =   37
         Top             =   315
         Visible         =   0   'False
         Width           =   9045
         Begin VB.TextBox txtUName 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   39
            Top             =   1875
            Width           =   2805
         End
         Begin VB.TextBox txtUID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   1530
            Width           =   2820
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "University :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Index           =   12
            Left            =   1440
            TabIndex        =   41
            Top             =   1905
            Width           =   1500
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "UniversityID  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   8
            Left            =   1395
            TabIndex        =   40
            Top             =   1530
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3975
         Left            =   45
         TabIndex        =   29
         Top             =   315
         Width           =   9045
         Begin VB.TextBox txtCity_dist 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   1170
            Width           =   1785
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Select Branch Code"
            Top             =   1125
            Width           =   510
         End
         Begin VB.TextBox txtCityID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1530
            Width           =   2820
         End
         Begin VB.TextBox txtCity_disId 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1170
            Width           =   1020
         End
         Begin VB.TextBox txtCityName 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   30
            Top             =   1875
            Width           =   2805
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "City/Town-Id  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   7
            Left            =   1350
            TabIndex        =   36
            Top             =   1530
            Width           =   1200
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "District :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   6
            Left            =   1350
            TabIndex        =   35
            Top             =   1170
            Width           =   1200
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "City/Town  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Index           =   5
            Left            =   1350
            TabIndex        =   34
            Top             =   1905
            Width           =   1590
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3930
         Left            =   -74955
         TabIndex        =   21
         Top             =   360
         Width           =   9045
         Begin VB.TextBox txtDisState 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            TabIndex        =   66
            Top             =   1170
            Width           =   1830
         End
         Begin VB.TextBox txtDis_Name 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   25
            Top             =   1875
            Width           =   2580
         End
         Begin VB.TextBox txtDis_StateId 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   4365
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1170
            Width           =   750
         End
         Begin VB.TextBox txtDis_DistID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1530
            Width           =   2595
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   5175
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Select Branch Code"
            Top             =   1125
            Width           =   510
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "District :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Index           =   10
            Left            =   1530
            TabIndex        =   28
            Top             =   1905
            Width           =   1410
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "State :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   4
            Left            =   1530
            TabIndex        =   27
            Top             =   1170
            Width           =   1020
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "District-Id  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   3
            Left            =   1530
            TabIndex        =   26
            Top             =   1530
            Width           =   1020
         End
      End
      Begin VB.Frame Container 
         Height          =   3930
         Left            =   -74955
         TabIndex        =   12
         Top             =   315
         Width           =   9045
         Begin VB.TextBox txtCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1530
            Width           =   2595
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   13
            Top             =   1875
            Width           =   2580
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Country-Id  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   1
            Left            =   1530
            TabIndex        =   16
            Top             =   1530
            Width           =   1425
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " *Country :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Index           =   11
            Left            =   1440
            TabIndex        =   15
            Top             =   1905
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3930
         Left            =   -74955
         TabIndex        =   7
         Top             =   315
         Width           =   9045
         Begin VB.TextBox txtCountry 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            TabIndex        =   65
            Top             =   1170
            Width           =   1920
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   5310
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Select Branch Code"
            Top             =   1125
            Width           =   510
         End
         Begin VB.TextBox txtSt_StateId 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1530
            Width           =   2685
         End
         Begin VB.TextBox txtSt_CountryID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Pub_code"
            Height          =   285
            Left            =   4455
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1170
            Width           =   750
         End
         Begin VB.TextBox txtSt_State 
            Appearance      =   0  'Flat
            DataField       =   "Name"
            Height          =   285
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   8
            Top             =   1875
            Width           =   2670
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "State-Id  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   9
            Left            =   1530
            TabIndex        =   18
            Top             =   1530
            Width           =   1020
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Country  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   270
            Index           =   2
            Left            =   1530
            TabIndex        =   11
            Top             =   1170
            Width           =   1425
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "State :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Index           =   0
            Left            =   1530
            TabIndex        =   10
            Top             =   1905
            Width           =   1410
         End
      End
   End
End
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit As Boolean
Dim rs_map As New ADODB.Recordset
Private Sub cmdAdd_1_Click()




If SSTab1.Tab = 0 Then

    txtCode.Text = MaxSNo_New("Country", "Countryid", "Country")
    txtName = ""
    txtName.SetFocus

ElseIf SSTab1.Tab = 1 Then

    txtSt_StateId = MaxSNo_New("state", "stateid", "state")
    
    txtCountry.SetFocus
    txtSt_State = ""
    'txtSt_State.SetFocus
     
ElseIf SSTab1.Tab = 2 Then

    txtDis_DistID = MaxSNo_New("District", "Districtid", "District")
    txtDis_Name = ""
    
    txtDisState.SetFocus
    
    'txtDis_Name.SetFocus

ElseIf SSTab1.Tab = 3 Then

    txtCityID = MaxSNo_New("city", "cityid", "city")
    txtCityName = ""
    txtCity_dist.SetFocus

ElseIf SSTab1.Tab = 4 Then

    txtUID = MaxSNo_New("University", "Universityid", "University")
    txtUName = ""
    'txtUName.SetFocus

ElseIf SSTab1.Tab = 5 Then

    txtDepartmentId = MaxSNo_New("Department", "Departmentid", "Department")
    txtDepartment = ""
    'txtDepartment.SetFocus

ElseIf SSTab1.Tab = 6 Then

    txtBookId = MaxSNo_New("BookType", "BookTypeid", "BookType")
    txtBookType = ""
    'xtBookType.SetFocus

ElseIf SSTab1.Tab = 7 Then

    txtAuthId = MaxSNo_New("Auther", "Autherid", "Auther")
    txtAuther = ""
    txtAddress1 = ""
    txtAddress2 = ""
    txtCityID = ""
    txtPhone1 = ""
    txtphone2 = ""
    txtpin = ""
    txtCity = ""
    txtDist = ""
    txtCityID = ""
    txtState = ""
    txtCityID1 = ""
    txtPan = ""
     
    'txtAuther.SetFocus

End If


cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdEdit_4.Enabled = True
edit = False

Me.Caption = ""

End Sub

Private Sub cmdAdd_Click()
If SSTab1.Tab = 1 Then
   tblNo = 2
   frmSearchItem.Show
   'popuplist2 "select Country,CountryId from Country", CON
End If
End Sub
Private Sub cmdAdd_GotFocus()
   
   If PopUpValue1 <> "" Then
      
   If SSTab1.Tab = 1 Then
       txtSt_CountryID.Tag = PopUpValue2
       txtCountry = PopUpValue2
       txtSt_CountryID.Text = PopUpValue1
       Call cmdAdd_1_Click
       txtSt_State.SetFocus
   End If
      
   End If
   
   
      PopUpValue1 = ""
      PopUpValue2 = ""
   
End Sub
Private Sub cmdDelete_3_Click()
On Error GoTo err:

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'=============================================================
If MsgBox("Want To Delete", vbQuestion + vbYesNo) = vbNo Then Exit Sub

If SSTab1.Tab = 0 Then

CON.BeginTrans
CON.Execute "DELETE FROM  [Country] WHERE countryid='" & txtCode & "'"
CON.CommitTrans

cmdAdd_1_Click

ElseIf SSTab1.Tab = 1 Then

CON.BeginTrans
CON.Execute "DELETE FROM  [state] WHERE stateid='" & txtSt_StateId & "'"
CON.CommitTrans

cmdAdd_1_Click

txtSt_State.SetFocus

ElseIf SSTab1.Tab = 2 Then

CON.BeginTrans
CON.Execute "DELETE FROM  [District] WHERE Districtid='" & txtDis_DistID & "'"
CON.CommitTrans

cmdAdd_1_Click
txtDisState.SetFocus

ElseIf SSTab1.Tab = 3 Then

CON.BeginTrans
CON.Execute "DELETE FROM  [city] WHERE cityid='" & txtCityID & "'"
CON.CommitTrans

txtCity_dist.SetFocus
cmdAdd_1_Click

ElseIf SSTab1.Tab = 4 Then

CON.BeginTrans
CON.Execute "DELETE FROM  [University] WHERE Universityid='" & txtUID & "'"
CON.CommitTrans

txtUID.SetFocus
cmdAdd_1_Click

ElseIf SSTab1.Tab = 5 Then

CON.BeginTrans
CON.Execute "DELETE FROM  [Department] WHERE Departmentid='" & txtDepartmentId & "'"
CON.CommitTrans


txtDepartmentId.SetFocus
cmdAdd_1_Click

ElseIf SSTab1.Tab = 6 Then

CON.BeginTrans
CON.Execute "DELETE FROM  [BookType] WHERE BookTypeid='" & txtBookId & "'"
CON.CommitTrans

txtBookId.SetFocus
cmdAdd_1_Click

ElseIf SSTab1.Tab = 7 Then

CON.BeginTrans
CON.Execute "DELETE FROM  [Auther] WHERE Autherid='" & txtAuthId & "'"
CON.CommitTrans

txtAuthId.SetFocus
cmdAdd_1_Click
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'=============================================================


'message Me, "d"



Exit Sub
err:

CON.RollbackTrans
MsgBox "" & err.DESCRIPTION



End Sub
Private Sub cmdEdit_4_Click()
    edit = True
    cmdEdit_4.Enabled = False
    cmdSave_2.Enabled = True
    cmdSave_2.SetFocus
    
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub
Private Sub cmdSave_2_Click()

On Error GoTo err:

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'=============================================================

If SSTab1.Tab = 0 Then


If txtName = "" Then
 MsgBox "Plz. Country Name ...", vbCritical
 txtName.SetFocus
 Exit Sub
End If


If edit = False Then

txtCode.Text = MaxSNo_New("Country", "Countryid", "Country")
CON.BeginTrans
CON.Execute "INSERT INTO  [Country]" & _
           "([CountryID]" & _
           ",[Country])" & _
     "Values" & _
           "('" & txtCode & "'," & _
           "'" & txtName & "')"
CON.CommitTrans

Else


CON.BeginTrans
CON.Execute "Update  [Country]" & _
           "SET Country = '" & txtName & "'" & _
           " where CountryID='" & txtCode & "'"

CON.CommitTrans



End If

cmdAdd_1_Click
txtName.SetFocus

ElseIf SSTab1.Tab = 1 Then

'============================================== State Saving Code===============================

If txtSt_CountryID = "" Then
   MsgBox "Plz. Select Country Name ...", vbCritical
   txtSt_CountryID.SetFocus
   Exit Sub
End If

If txtSt_State = "" Then
   MsgBox "Plz. Enter Sate Name ...", vbCritical
   txtSt_State.SetFocus
   Exit Sub
End If




If edit = False Then


txtSt_StateId = MaxSNo_New("state", "stateid", "state")
CON.BeginTrans
CON.Execute "INSERT INTO  [State]" & _
           "([CountryID]" & _
           ",[StateID]" & _
           ",[State]" & _
           ",[StateKey])" & _
     "Values" & _
           "('" & txtSt_CountryID & "'," & _
           "'" & txtSt_StateId & "'," & _
           "'" & txtSt_State & "'," & _
           "'" & txtSt_State & " ~ " & txtSt_CountryID.Tag & "')"
CON.CommitTrans

Else

CON.BeginTrans
CON.Execute "Update  [State]" & _
           "SET [CountryID] = '" & txtSt_CountryID.Text & "'" & _
           ",[State]='" & txtSt_State & "'" & _
           ",[StateKey]='" & txtSt_State & " ~ " & txtSt_CountryID.Tag & "' where StateID='" & txtSt_StateId & "'"

CON.CommitTrans

End If
'============================================== end Code======================================

cmdAdd_1_Click
txtSt_State.SetFocus

ElseIf SSTab1.Tab = 2 Then


If txtDis_StateId = "" Then
   MsgBox "Plz. Select Sate ...", vbCritical
   txtDis_StateId.SetFocus
   Exit Sub
End If

If txtDis_Name = "" Then
   MsgBox "Plz. Enter District Name ...", vbCritical
   txtDis_Name.SetFocus
   Exit Sub
End If


'============================================== District Saving Code===============================

If edit = False Then


txtDis_DistID = MaxSNo_New("District", "Districtid", "District")
CON.BeginTrans
CON.Execute "INSERT INTO  [District]" & _
           "([stateID]" & _
           ",[DistrictID]" & _
           ",[District]" & _
           ",[DistrictKey])" & _
     "Values" & _
           "('" & txtDis_StateId & "'," & _
           "'" & txtDis_DistID & "'," & _
           "'" & txtDis_Name & "'," & _
           "'" & txtDis_Name & " ~ " & txtDis_StateId.Tag & "')"
CON.CommitTrans

Else

CON.BeginTrans
CON.Execute "Update  [District]" & _
           "SET [stateID] = '" & txtDis_StateId & "'" & _
           ",[District]='" & txtDis_Name & "'" & _
           ",[DistrictKey]='" & txtDis_Name & " ~ " & txtDis_StateId.Tag & "' where DistrictID='" & txtDis_DistID & "'"

CON.CommitTrans

End If

cmdAdd_1_Click
txtDis_Name.SetFocus

'============================================== end Code===============================

ElseIf SSTab1.Tab = 3 Then


If txtCity_disId = "" Then
   MsgBox "Plz. Select District ...", vbCritical
   txtDis_DistID.SetFocus
   Exit Sub
End If

If txtCityName = "" Then
   MsgBox "Plz. Enter City Name ...", vbCritical
   txtCityName.SetFocus
   Exit Sub
End If


'============================================== City Saving Code===============================

If edit = False Then


txtUID = MaxSNo_New("University", "Universityid", "University")
CON.BeginTrans
CON.Execute "INSERT INTO  [City]" & _
           "([DistrictID]" & _
           ",[CityID]" & _
           ",[City]" & _
           ",[CityKey])" & _
     "Values" & _
           "('" & txtCity_disId & "'," & _
           "'" & txtCityID & "'," & _
           "'" & txtCityName & "'," & _
           "'" & txtCityName & " ~ " & txtCity_disId.Tag & "')"
CON.CommitTrans

Else

CON.BeginTrans
CON.Execute "Update  [City]" & _
           "SET [DistrictID] = '" & txtCity_disId & "'" & _
           ",[city]='" & txtCityName & "'" & _
           ",[cityKey]='" & txtCityName & " ~ " & txtCity_disId.Tag & "' where cityID='" & txtCityID & "'"

CON.CommitTrans

End If

cmdAdd_1_Click
'============================================== end Code===============================

ElseIf SSTab1.Tab = 4 Then


If txtUName = "" Then
 MsgBox "Plz. Enter University Name ...", vbCritical
 txtUName.SetFocus
 Exit Sub
End If

If edit = False Then

txtUID.Text = MaxSNo_New("University", "Universityid", "University")
CON.BeginTrans
CON.Execute "INSERT INTO  [University]" & _
           "([UniversityID]" & _
           ",[University])" & _
     "Values" & _
           "('" & txtUID & "'," & _
           "'" & txtUName & "')"
CON.CommitTrans

Else


CON.BeginTrans
CON.Execute "Update  [University]" & _
           "SET University = '" & txtUName & "'" & _
           " where UniversityID='" & txtUID & "'"

CON.CommitTrans



End If

cmdAdd_1_Click

ElseIf SSTab1.Tab = 5 Then


If txtDepartment = "" Then
 MsgBox "Plz. Enter Department Name ...", vbCritical
 txtDepartment.SetFocus
 Exit Sub
End If

If edit = False Then

txtDepartmentId.Text = MaxSNo_New("Department", "Departmentid", "Department")
CON.BeginTrans
CON.Execute "INSERT INTO  [Department]" & _
           "([DepartmentID]" & _
           ",[Department])" & _
     "Values" & _
           "('" & txtDepartmentId & "'," & _
           "'" & txtDepartment & "')"
CON.CommitTrans

Else


CON.BeginTrans
CON.Execute "Update  [Department]" & _
           "SET Department = '" & txtDepartment & "'" & _
           " where DepartmentID='" & txtDepartmentId & "'"

CON.CommitTrans

End If

cmdAdd_1_Click

ElseIf SSTab1.Tab = 6 Then


If txtBookType = "" Then
    MsgBox "Plz. Enter BookType ...", vbCritical
    txtBookType.SetFocus
    Exit Sub
End If

If edit = False Then

txtBookId.Text = MaxSNo_New("BookType", "BookTypeid", "BookType")

CON.BeginTrans
CON.Execute "INSERT INTO  [BookType]" & _
           "([BookTypeID]" & _
           ",[BookType])" & _
     "Values" & _
           "('" & txtBookId & "'," & _
           "'" & txtBookType & "')"
CON.CommitTrans


Else


CON.BeginTrans
CON.Execute "Update  [BookType]" & _
           "SET BookType = '" & txtBookType & "'" & _
           " where BookTypeID='" & txtBookId & "'"
CON.CommitTrans
End If

cmdAdd_1_Click

ElseIf SSTab1.Tab = 7 Then



If txtAuther = "" Then
    MsgBox "Plz. Enter Auther ...", vbCritical
    txtAuther.SetFocus
    Exit Sub
End If

If edit = False Then

txtAuthId = MaxSNo_New("Auther", "Autherid", "Auther")

CON.BeginTrans
CON.Execute "INSERT INTO  [Auther]" & _
           "([AutherID]" & _
           ",[Auther]" & _
           ",[Add1]" & _
           ",[Add2]" & _
           ",[CityID]" & _
           ",[Phone]" & _
           ",[Pin],[Pan])" & _
     "Values" & _
           "('" & txtAuthId & "'," & _
           "'" & txtAuther & "'," & _
           "'" & txtAddress1 & "'," & _
           "'" & txtAddress2 & "'," & _
           "'" & txtCityID1 & "'," & _
           "'" & txtPhone1 & "'," & _
           "'" & txtpin.Text & "','" & txtPan.Text & "')"
CON.CommitTrans


Else


CON.BeginTrans
CON.Execute "Update  [Auther]" & _
           "SET Auther = '" & txtAuther & "',[Add1]='" & txtAddress1.Text & "',[Add2]='" & txtAddress2.Text & "',phone='" & txtPhone1 & "',PIN='" & txtpin.Text & "'," & _
           "pan='" & txtPan.Text & "' where AutherID='" & txtAuthId & "'"
CON.CommitTrans


End If




cmdAdd_1_Click


End If




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'message Me, "s"
'cmdSave_2.Enabled = False
'edit = False

'=============================================================


Exit Sub
err:

CON.RollbackTrans
MsgBox "" & err.DESCRIPTION
           
End Sub
Sub search_Data()
If SSTab1.Tab = 0 Then
   popuplist10 "select Country,CountryId from Country order by Country", CON
ElseIf SSTab1.Tab = 1 Then
   'popuplist2 "select State,StateKey,StateID,CountryID from State order by State", CON
   tblNo = 7
   frmSearchItem.Show

ElseIf SSTab1.Tab = 2 Then
   
   tblNo = 4
   frmSearchItem.Show
   
ElseIf SSTab1.Tab = 3 Then
   'popuplist2 "select City,CityKey,CityID,DistrictID from City order by city", CON
   tblNo = 6
   frmSearchItem.Show

   
'ElseIf SSTab1.Tab = 4 Then
'   popuplist2 "select University,UniversityId from University order by University", CON
'ElseIf SSTab1.Tab = 5 Then
'   popuplist2 "select Department,DepartmentId from Department order by Department", CON
'ElseIf SSTab1.Tab = 6 Then
'   popuplist2 "select BookType,BookTypeId from BookType order by BookType", CON
'ElseIf SSTab1.Tab = 7 Then
'   popuplist2 "select Auther,AutherId from Auther order by Auther", CON
End If

End Sub

Private Sub cmdSearch_Click()
If SSTab1.Tab = 0 Then
   popuplist10 "select Country,CountryId from Country order by Country", CON
ElseIf SSTab1.Tab = 1 Then
   'popuplist2 "select State,StateKey,StateID,CountryID from State order by State", CON
   tblNo = 7
   frmSearchItem.Show

ElseIf SSTab1.Tab = 2 Then
   
   tblNo = 4
   frmSearchItem.Show
   
ElseIf SSTab1.Tab = 3 Then
   
   'popuplist2 "select City,CityKey,CityID,DistrictID from City order by city", CON
   tblNo = 6
   frmSearchItem.Show

   
'ElseIf SSTab1.Tab = 4 Then
'   popuplist10 "select University,UniversityId from University order by University", CON
'ElseIf SSTab1.Tab = 5 Then
'   popuplist2 "select Department,DepartmentId from Department order by Department", CON
'ElseIf SSTab1.Tab = 6 Then
'   popuplist2 "select BookType,BookTypeId from BookType order by BookType", CON
'ElseIf SSTab1.Tab = 7 Then
'   'popuplist2 "select Auther,AutherId from Auther order by Auther", CON
'   tblNo = 15
'   frmSearchItem.Show

End If

End Sub
Private Sub cmdSearch_GotFocus()

On Error GoTo aa1:

If SSTab1.Tab = 0 Then

If PopUpValue1 <> "" Then
   txtCode = PopUpValue2
   txtName = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
End If

ElseIf SSTab1.Tab = 1 Then

If PopUpValue1 <> "" Then
   
   
   cmdEdit_4_Click
   
   
   txtSt_State = PopUpValue2
   txtSt_StateId = PopUpValue1
   
   'txtSt_CountryID.Text = PopUpValue4
   'txtSt_CountryID.Tag = Country_State_Dist_City(PopUpValue4, "country")
   'txtCountry = txtSt_CountryID.Tag
   rs_map.MoveFirst
    rs_map.Find "[StateID]='" & txtSt_StateId & "'"
    If rs_map.EOF = False Then
       txtSt_CountryID.Text = rs_map![CountryID]
       txtSt_CountryID.Tag = rs_map![country]
       txtCountry = rs_map![country]
    End If
      
      
   PopUpValue1 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
End If


ElseIf SSTab1.Tab = 2 Then

If PopUpValue1 <> "" Then
   
   
   cmdEdit_4_Click
   
   'rs_map.Requery
   rs_map.MoveFirst
   rs_map.Find "DistrictID='" & PopUpValue1 & "'"
   txtDis_Name = rs_map!District
   txtDis_DistID = rs_map!DistrictID
   txtDis_StateId.Text = rs_map!StateID
   txtDis_StateId.Tag = rs_map![State]
   txtDisState = rs_map![State]  ' rs_map![State.StateID]
      
   PopUpValue1 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
   
End If

ElseIf SSTab1.Tab = 3 Then

If PopUpValue1 <> "" Then
   
   
   cmdEdit_4_Click
   
   'rs_map.Requery
   rs_map.MoveFirst
   rs_map.Find "[CityID]='" & PopUpValue1 & "'"
   txtCityName = rs_map!CITY
   txtCityID = PopUpValue1
   txtCity_disId.Text = rs_map![DistrictID]
   txtCity_disId.Tag = rs_map![District]
   txtCity_dist.Text = rs_map![District]
      
   PopUpValue1 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
End If

ElseIf SSTab1.Tab = 4 Then

If PopUpValue1 <> "" Then
   txtUID = PopUpValue2
   txtUName = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
   cmdSave_2.Enabled = False
End If

ElseIf SSTab1.Tab = 5 Then

If PopUpValue1 <> "" Then
   txtDepartmentId = PopUpValue2
   txtDepartment = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
   cmdSave_2.Enabled = False
End If

ElseIf SSTab1.Tab = 6 Then

If PopUpValue1 <> "" Then
   txtBookId = PopUpValue2
   txtBookType = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
   cmdSave_2.Enabled = False
End If

ElseIf SSTab1.Tab = 7 Then

If PopUpValue1 <> "" Then
   txtAuthId = PopUpValue1
   txtAuther = PopUpValue2
   
   If rs.State = 1 Then rs.Close
   rs.Open "select * from [Auther] where [AutherID]='" & PopUpValue1 & "'", CON
   If rs.EOF = False Then
    txtAddress1 = rs![add1] & ""
    txtAddress2 = rs![add2] & ""
    txtCityID1 = rs![CityID] & ""
    txtPhone1 = rs![PHONE] & ""
    txtpin.Text = rs![pin] & ""
    txtPan = rs![pan] & ""
   End If
   rs_map.MoveFirst
   rs_map.Find "[CityID]='" & rs![CityID] & "'"
   If rs_map.EOF = False Then
      txtCity = rs_map!CITY & ""
      txtDist = rs_map!District & ""
      txtState = rs_map!State & ""
   End If

   
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   cmdSave_2.Enabled = False
End If


End If





cmdDelete_3.Enabled = True
cmdEdit_4.Enabled = True
'cmdSave_2.Enabled = False

PopUpValue2 = ""
PopUpValue1 = ""


Exit Sub
aa1:




End Sub
Sub searchGotFocus()
If SSTab1.Tab = 0 Then

If PopUpValue1 <> "" Then
   txtCode = PopUpValue2
   txtName = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
End If

ElseIf SSTab1.Tab = 1 Then

If PopUpValue1 <> "" Then
   
   
   cmdEdit_4_Click
   
   
   txtSt_State = PopUpValue2
   txtSt_StateId = PopUpValue1
   
   'txtSt_CountryID.Text = PopUpValue4
   'txtSt_CountryID.Tag = Country_State_Dist_City(PopUpValue4, "country")
   'txtCountry = txtSt_CountryID.Tag
   rs_map.Requery
   rs_map.MoveFirst
   rs_map.Find "[State.StateID]='" & txtSt_StateId & "'"
   If rs_map.EOF = False Then
      txtSt_CountryID.Text = rs_map![Country.CountryID]
      txtSt_CountryID.Tag = rs_map![country]
      txtCountry = rs_map![country]
   End If
      
      
   PopUpValue1 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
End If


ElseIf SSTab1.Tab = 2 Then

If PopUpValue1 <> "" Then
   
   
   cmdEdit_4_Click
   
   rs_map.Requery
   rs_map.MoveFirst
   rs_map.Find "[District.DistrictID]='" & PopUpValue1 & "'"
   txtDis_Name = rs_map!District
   txtDis_DistID = rs_map![District.DistrictID]
   txtDis_StateId.Text = rs_map![State.StateID]
   txtDis_StateId.Tag = rs_map![State]
   txtDisState = rs_map![State]  ' rs_map![State.StateID]
      
   PopUpValue1 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
End If

ElseIf SSTab1.Tab = 3 Then

If PopUpValue1 <> "" Then
   
   
   cmdEdit_4_Click
   
   
   rs_map.MoveFirst
   rs_map.Find "[CityID]='" & PopUpValue1 & "'"
   txtCityName = rs_map!CITY
   txtCityID = PopUpValue1
   txtCity_disId.Text = rs_map![District.DistrictID]
   txtCity_disId.Tag = rs_map![District]
   txtCity_dist.Text = rs_map![District]
      
   PopUpValue1 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
End If

ElseIf SSTab1.Tab = 4 Then

If PopUpValue1 <> "" Then
   txtUID = PopUpValue2
   txtUName = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
End If

ElseIf SSTab1.Tab = 5 Then

If PopUpValue1 <> "" Then
   txtDepartmentId = PopUpValue2
   txtDepartment = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
End If

ElseIf SSTab1.Tab = 6 Then

If PopUpValue1 <> "" Then
   txtBookId = PopUpValue2
   txtBookType = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
End If

ElseIf SSTab1.Tab = 7 Then

If PopUpValue1 <> "" Then
   txtAuthId = PopUpValue2
   txtAuther = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
End If


End If





cmdDelete_3.Enabled = True
cmdEdit_4.Enabled = True
'cmdSave_2.Enabled = False

PopUpValue2 = ""
PopUpValue1 = ""

End Sub

Private Sub Command1_Click()

'If SSTab1.Tab = 2 Then
'   popuplist2 "select State,StateId,StateKey from State order by State", CON
'End If

If SSTab1.Tab = 2 Then
   tblNo = 7
   frmSearchItem.Show
   'popuplist2 "select Country,CountryId from Country", CON
End If


End Sub
Private Sub Command1_GotFocus()
   
   If PopUpValue1 <> "" Then
      
   If SSTab1.Tab = 2 Then
       txtDis_StateId.Tag = PopUpValue2
       txtDis_StateId.Text = PopUpValue1
       txtDisState = PopUpValue2
   End If
      
   End If
   
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   
End Sub

Private Sub Command2_Click()

'If SSTab1.Tab = 3 Then
'   popuplist2 "select District,DistrictId,DistrictKey from District order by District", CON
'End If

If SSTab1.Tab = 3 Then
   tblNo = 4
   frmSearchItem.Show
End If


End Sub
Private Sub Command2_GotFocus()

  If PopUpValue1 <> "" Then
      
   If SSTab1.Tab = 3 Then
       txtCity_disId.Tag = PopUpValue2
       txtCity_disId.Text = PopUpValue1
       txtCity_dist.Text = PopUpValue2
       txtCityName.SetFocus
   End If
      
   End If
   
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""

End Sub
Private Sub Form_Load()
txtCode.Text = MaxSNo_New("Country", "Countryid", "Country")
'frmBack Me
'formDisplaySetting Me

If rs_map.State = 1 Then rs_map.Close
rs_map.Open "select * from qrymap", CON, adOpenKeyset, adLockReadOnly


End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  cmdAdd_1_Click
End Sub


Private Sub txtAddress1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAddress2.SetFocus
End Sub

Private Sub txtAddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCity.SetFocus
End Sub

Private Sub txtAuther_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtAddress1.SetFocus
End Sub

Private Sub txtBookType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdSave_2_Click
End Sub

Private Sub txtCity_dist_GotFocus()
  If PopUpValue1 <> "" Then
      
   If SSTab1.Tab = 3 Then
       txtCity_disId.Tag = PopUpValue2
       txtCity_disId.Text = PopUpValue1
       txtCity_dist.Text = PopUpValue2
       HIT
       txtCityName.SetFocus
   End If
      
   End If
   
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""

End Sub

Private Sub txtCity_dist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCityName.SetFocus
End Sub

Private Sub txtCity_dist_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Exit Sub
If KeyCode = 13 Then Exit Sub

If SSTab1.Tab = 3 Then
   
   tblNo = 4
   frmSearchItem.Show
End If

End Sub

Private Sub txtcity_GotFocus()
If PopUpValue1 <> "" Then
   
    txtCity = PopUpValue2
    txtCityID1 = PopUpValue1
    
    If rs.State = 1 Then rs.Close
    rs.Open "select [District],[State] FROM  [CityView] " & _
    "where [CityID]='" & PopUpValue1 & "'", CON
    If rs.EOF = False Then
       txtDist = rs(0)
       txtState = rs(1)
    End If
    
    txtpin.SetFocus
End If

PopUpValue1 = ""
PopUpValue2 = ""

End Sub

Private Sub txtcity_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then Exit Sub
   If KeyCode = 13 Then Exit Sub
   
   HIT
   tblNo = 6
   frmSearchItem.Show
End Sub

Private Sub txtCityName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdSave_2_Click
End Sub

Private Sub txtCountry_GotFocus()
   
   If PopUpValue1 <> "" Then
      
   If SSTab1.Tab = 1 Then
       txtSt_CountryID.Tag = PopUpValue2
       txtCountry = PopUpValue2
       txtSt_CountryID.Text = PopUpValue1
       Call cmdAdd_1_Click
       HIT
       txtSt_State.SetFocus
   End If
      
   End If
   
   
      PopUpValue1 = ""
      PopUpValue2 = ""

End Sub




Private Sub txtCountry_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Exit Sub

If SSTab1.Tab = 1 Then
   tblNo = 2
   popupvalue5 = txtCountry
   frmSearchItem.Show
   'popuplist2 "select Country,CountryId from Country", CON
End If
End Sub
Private Sub txtDepartment_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdSave_2_Click
End Sub

Private Sub txtDis_Name_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdSave_2_Click
End Sub

Private Sub txtDisState_GotFocus()
   

   If PopUpValue1 <> "" Then
      
   If SSTab1.Tab = 2 Then
       txtDis_StateId.Tag = PopUpValue2
       txtDis_StateId.Text = PopUpValue1
       txtDisState = PopUpValue2
       HIT
       txtDis_Name.SetFocus
   End If
      
   End If
   
   
   PopUpValue1 = ""
   PopUpValue2 = ""

End Sub
Private Sub txtDisState_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Exit Sub

If SSTab1.Tab = 2 Then
   HIT
   popupvalue5 = txtDisState
   tblNo = 7
   frmSearchItem.Show
End If

End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdSave_2_Click
End Sub
Private Sub txtPan_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdSave_2_Click
End Sub

Private Sub txtPhone1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPan.SetFocus
End Sub

Private Sub txtpin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then txtPhone1.SetFocus
End Sub

Private Sub txtSt_State_GotFocus()
'  searchGotFocus
End Sub

Private Sub txtSt_State_KeyDown(KeyCode As Integer, Shift As Integer)
'  search_Data
End Sub
Private Sub txtSt_State_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSave_2_Click
End Sub

Private Sub txtUName_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdSave_2_Click
End Sub
