VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form setup 
   ClientHeight    =   9684
   ClientLeft      =   276
   ClientTop       =   1428
   ClientWidth     =   9660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "setup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9684
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "check Trial"
      Height          =   504
      Left            =   720
      TabIndex        =   75
      Top             =   7236
      Width           =   1212
   End
   Begin VB.CommandButton Command6 
      Caption         =   "check Missing Voucher"
      Height          =   492
      Left            =   4968
      TabIndex        =   74
      Top             =   7272
      Visible         =   0   'False
      Width           =   408
   End
   Begin VB.CommandButton cmdCreateMail 
      Caption         =   "Create Total Book Name School Wise"
      Height          =   435
      Left            =   7560
      TabIndex        =   73
      Top             =   7704
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Map App.Form to invoice"
      Height          =   570
      Left            =   5355
      TabIndex        =   72
      Top             =   7290
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdtmpClear 
      Caption         =   "Clear Tmp"
      Height          =   570
      Left            =   6435
      TabIndex        =   71
      Top             =   7290
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton Command4 
      Caption         =   "update TOD"
      Height          =   456
      Left            =   7245
      TabIndex        =   70
      Top             =   7668
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "update free Table  from main table"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7245
      TabIndex        =   57
      Top             =   7290
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Frame Frame2 
      Height          =   6795
      Left            =   240
      TabIndex        =   26
      Top             =   360
      Width           =   9015
      Begin VB.Frame Frame4 
         Caption         =   $"setup.frx":000C
         Height          =   915
         Left            =   1800
         TabIndex        =   64
         Top             =   5760
         Width           =   7035
         Begin MSMask.MaskEdBox txtFromSale_nyr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   345
            Left            =   660
            TabIndex        =   21
            Top             =   480
            Width           =   1035
            _ExtentX        =   1820
            _ExtentY        =   614
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtToSale_nyr 
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
            Left            =   2040
            TabIndex        =   22
            Top             =   480
            Width           =   1035
            _ExtentX        =   1820
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFromSaleRet_nyr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   345
            Left            =   4260
            TabIndex        =   23
            Top             =   420
            Width           =   1035
            _ExtentX        =   1820
            _ExtentY        =   614
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtToSaleRet_nyr 
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
            Left            =   5640
            TabIndex        =   24
            Top             =   420
            Width           =   1035
            _ExtentX        =   1820
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "From "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   68
            Top             =   540
            Width           =   480
         End
         Begin VB.Label Label13 
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1740
            TabIndex        =   67
            Top             =   480
            Width           =   345
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "From "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3660
            TabIndex        =   66
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label11 
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5340
            TabIndex        =   65
            Top             =   420
            Width           =   345
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   $"setup.frx":009A
         Height          =   915
         Left            =   1800
         TabIndex        =   58
         Top             =   4680
         Width           =   7035
         Begin MSMask.MaskEdBox txtFromSale_cyr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   345
            Left            =   660
            TabIndex        =   17
            Top             =   480
            Width           =   1035
            _ExtentX        =   1820
            _ExtentY        =   614
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtToSale_cyr 
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
            Left            =   2040
            TabIndex        =   18
            Top             =   480
            Width           =   1035
            _ExtentX        =   1820
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFromSaleRet_cyr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   345
            Left            =   4260
            TabIndex        =   19
            Top             =   420
            Width           =   1035
            _ExtentX        =   1820
            _ExtentY        =   614
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtToSaleRet_cyr 
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
            Left            =   5640
            TabIndex        =   20
            Top             =   420
            Width           =   1035
            _ExtentX        =   1820
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5340
            TabIndex        =   63
            Top             =   420
            Width           =   345
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "From "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3660
            TabIndex        =   62
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label7 
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1740
            TabIndex        =   60
            Top             =   480
            Width           =   345
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "From "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   59
            Top             =   540
            Width           =   480
         End
      End
      Begin VB.TextBox txtcst 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5820
         MaxLength       =   244
         TabIndex        =   13
         Top             =   2760
         Width           =   3045
      End
      Begin VB.TextBox txtuptt 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1830
         MaxLength       =   244
         TabIndex        =   12
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox txtbankadvice 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   11
         Top             =   2400
         Width           =   1995
      End
      Begin VB.TextBox txtemail 
         DataField       =   "FOOTNOTE1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   8
         Top             =   1650
         Width           =   4845
      End
      Begin VB.TextBox txtfax 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6990
         MaxLength       =   244
         TabIndex        =   7
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox txtCname 
         DataField       =   "CLINIC_NAME"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   4905
      End
      Begin VB.TextBox add1 
         DataField       =   "ADDRESS1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   2
         Top             =   615
         Width           =   4935
      End
      Begin VB.TextBox add2 
         DataField       =   "ADDRESS2"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   3
         Top             =   945
         Width           =   3795
      End
      Begin VB.TextBox txtphone1 
         DataField       =   "PHONENO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   5
         Top             =   1275
         Width           =   2085
      End
      Begin VB.TextBox txtcourt 
         DataField       =   "FOOTNOTE1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1830
         MaxLength       =   244
         TabIndex        =   14
         Top             =   3150
         Width           =   7005
      End
      Begin VB.TextBox txtrem1 
         DataField       =   "FOOTNOTE2"
         DataSource      =   "Data1"
         Height          =   645
         Left            =   1830
         MaxLength       =   244
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Top             =   3525
         Width           =   6975
      End
      Begin VB.TextBox txtrem2 
         DataField       =   "FOOTNOTE3"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1830
         MaxLength       =   244
         TabIndex        =   16
         Top             =   4200
         Width           =   7035
      End
      Begin VB.TextBox txtphone2 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3960
         MaxLength       =   244
         TabIndex        =   6
         Top             =   1260
         Width           =   1995
      End
      Begin VB.TextBox CITY 
         DataField       =   "CITY"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6420
         MaxLength       =   244
         TabIndex        =   4
         Top             =   930
         Width           =   2415
      End
      Begin MSMask.MaskEdBox txtyarfrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   1845
         TabIndex        =   9
         Top             =   1980
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   614
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtyarto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   1980
         Width           =   1155
         _ExtentX        =   2032
         _ExtentY        =   656
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Next Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   69
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Current Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   61
         Top             =   5160
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Company  Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   41
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Foot Note Report :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   40
         Top             =   3180
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Phone No  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   660
         TabIndex        =   39
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   885
         TabIndex        =   38
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Remark 1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   750
         TabIndex        =   37
         Top             =   3570
         Width           =   945
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Remark 2 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   750
         TabIndex        =   36
         Top             =   4200
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "From Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   705
         TabIndex        =   35
         Top             =   2070
         Width           =   990
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1110
         TabIndex        =   34
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Bank Advice No :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   33
         Top             =   2460
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "City :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5790
         TabIndex        =   32
         Top             =   975
         Width           =   450
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Fax No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   6060
         TabIndex        =   31
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   30
         Top             =   2010
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "U.P.T.T. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   795
         TabIndex        =   29
         Top             =   2820
         Width           =   870
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "GSTIN :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   4890
         TabIndex        =   28
         Top             =   2820
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4425
      Left            =   840
      TabIndex        =   42
      Top             =   450
      Visible         =   0   'False
      Width           =   7965
      Begin VB.CheckBox cfooter 
         Caption         =   "Print Footer"
         Height          =   345
         Left            =   180
         TabIndex        =   51
         Top             =   2670
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox cheader 
         Caption         =   "Print Header"
         Height          =   345
         Left            =   180
         TabIndex        =   50
         Top             =   2220
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox footertext 
         Height          =   885
         Left            =   2220
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   2970
         Width           =   4785
      End
      Begin VB.TextBox headertext 
         Height          =   885
         Left            =   2220
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   1650
         Width           =   4785
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1290
         TabIndex        =   47
         Top             =   1710
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1290
         TabIndex        =   46
         Top             =   1320
         Width           =   555
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2550
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   3345
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   1905
         TabIndex        =   45
         Top             =   1710
         Width           =   240
         _ExtentX        =   445
         _ExtentY        =   487
         _Version        =   327681
         AutoBuddy       =   -1  'True
         OrigLeft        =   2610
         OrigTop         =   570
         OrigRight       =   2850
         OrigBottom      =   1245
         Max             =   35
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1845
         TabIndex        =   52
         Top             =   1320
         Width           =   240
         _ExtentX        =   445
         _ExtentY        =   508
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "cfooter"
         BuddyDispid     =   196647
         OrigLeft        =   2610
         OrigTop         =   570
         OrigRight       =   2850
         OrigBottom      =   1245
         Max             =   35
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label flabel 
         Caption         =   "Footer Text"
         Height          =   195
         Left            =   2220
         TabIndex        =   56
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label hlabel 
         Caption         =   "Header Text"
         Height          =   195
         Left            =   2220
         TabIndex        =   55
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Bottom Margin"
         Height          =   195
         Left            =   150
         TabIndex        =   54
         Top             =   1710
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Top Margin"
         Height          =   195
         Left            =   150
         TabIndex        =   53
         Top             =   1290
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Report Name :-"
         Height          =   315
         Left            =   1050
         TabIndex        =   44
         Top             =   270
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin MSComctlLib.TabStrip SStab1 
      Height          =   7230
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   9315
      _ExtentX        =   16425
      _ExtentY        =   12764
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Setup"
            Key             =   "Tab1"
            Object.Tag             =   "Tab1"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   465
      Left            =   3420
      TabIndex        =   27
      Top             =   7275
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   465
      Left            =   1995
      TabIndex        =   25
      Top             =   7275
      Width           =   1395
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   1668
      Left            =   72
      TabIndex        =   76
      Top             =   7920
      Visible         =   0   'False
      Width           =   7116
      _cx             =   12552
      _cy             =   2942
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"setup.frx":0128
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
Attribute VB_Name = "setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim RS As ADODB.Recordset
Dim next_Yrs As String
Private Sub cfooter_Click()
If cfooter.value = 1 Then
    flabel.Enabled = True
    footertext.Enabled = True
Else
    flabel.Enabled = False
    footertext.Enabled = False
End If
End Sub
Private Sub cheader_Click()

If cheader.value = 0 Then
   hlabel.Enabled = False
   headertext.Enabled = False
Else
   hlabel.Enabled = True
   headertext.Enabled = True
End If

End Sub
Private Sub cmdCreateMail_Click()
  Dim bkName_ As String
  Dim Email_ As String
  Dim rss As New ADODB.Recordset
  Dim con_ As New ADODB.Connection
  

  ''Set con_ = New ADODB.Connection
  ''Set RS = New ADODB.Recordset
  ''With con_
       ''.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + "tmpReport.mdb"
       ''.Open
  ''End With

  
  ''con.Execute "delete tmpBookSoldNameWithSchool"
  ''con.Execute "exec schoolWiseTotalBook"
  
  


  Dim fdate, tillDate
  fdate = "01/04/2018"
  tillDate = Format("30/07/2018", "dd/MM/yyyy")
  
  
  Dim str_date As String
  str_date = "(convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + fdate + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + tillDate + "',103))"
  
  
  
  If RS.State = 1 Then RS.close
  RS.Open "SELECT ScName,ScID FROM tmpSchoolWiseTotalBookSupply  group by ScName,ScID", con
  ssss = RS.RecordCount
  While RS.EOF = False
     bkName_ = ""
     Email_ = ""
     If rs1.State = 1 Then rs1.close
     rs1.Open "SELECT BOOKName FROM tmpSchoolWiseTotalBookSupply where (scid='" & RS!scid & "') group by BOOKName order by bookname", con, adOpenDynamic, adLockOptimistic
     While rs1.EOF = False
        If bkName_ = "" Then
           bkName_ = rs1!Bookname
        Else
           bkName_ = bkName_ & "," & rs1!Bookname
        End If
        rs1.MoveNext
     Wend
     

     
     If RS!scid <> "" Then
     If bkName_ <> "" Then
         
         con.Execute "update tmpSchoolqry_summary set TotalBookSupply='" & bkName_ & "' where ScID='" & RS!scid & "'"
         
     End If
     End If
   RS.MoveNext
  Wend
  
  
  MsgBox "ok"

End Sub

Private Sub cmdtmpClear_Click()

con.Execute "delete from tmpDDet"
con.Execute "delete from tmpDonnation"

MsgBox "tmp tbl Clear....", vbInformation


End Sub

Private Sub Command1_Click()

If SSTab1.Tabs(1).Key = "Tab1" Then

    If txtyarfrom.text = "__/__/____" Then
    MsgBox "Enter Session Period"
    txtyarfrom.SetFocus
    Exit Sub
    End If
    If txtyarto.text = "__/__/____" Then
    MsgBox "Enter Session Period"
    txtyarto.SetFocus
    Exit Sub
    End If
    
    If IsDate(txtyarfrom.text) = False Then
    MsgBox "Invalid Session Period"
    txtyarfrom.SetFocus
    Exit Sub
    End If
    
    If IsDate(txtyarto.text) = False Then
    MsgBox "Invalid Session Period"
    txtyarto.SetFocus
    Exit Sub
    End If
    
    If CStr(Year(CDate(Trim(txtyarfrom.text)))) <> Mid(main.session, 1, InStr(1, main.session, "-") - 1) Then
    MsgBox "Invalid Session Period"
    txtyarfrom.SetFocus
    Exit Sub
    End If
    If Year(Format(CDate(Trim(txtyarto.text)), "dd/mm/yy")) <> Year(CDate("01/01/" & Mid(main.session, InStr(1, main.session, "-") + 1))) Then
    MsgBox "Invalid Session Period"
    txtyarto.SetFocus
    Exit Sub
    End If
    
    
    
    
    Set RS = New ADODB.Recordset
    RS.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockOptimistic, adCmdText
    RS!cname = Me.txtCname.text
    RS!add1 = Me.add1.text
    RS!add2 = Me.add2.text
    RS!city = city.text
    RS!phone1 = Me.txtPhone1.text
    RS!phone2 = Me.txtphone2.text
    RS!FAX = Me.txtFax.text
    RS!yarfrom = txtyarfrom.text
    RS!yarto = txtyarto.text
    RS!email = txtEmail.text
    RS!COURT = txtcourt.text
    RS!rem1 = txtrem1.text
    RS!rem2 = txtrem2.text
    RS!bankadviceno = txtbankadvice.text
    RS!uptt = txtuptt.text
    RS!cst = txtcst.text
    RS.update




    Set RS = New ADODB.Recordset
    RS.Open "select * from turnOverDis where fyear='" & session & "' and Current_Next='current'", CCON, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       RS!fromdate = txtFromSale_cyr.text
       RS!todate = txtToSale_cyr.text
       
       RS!fromDateSRet = txtFromSaleRet_cyr.text
       RS!toDateSRet = txtToSaleRet_cyr.text
       RS.update
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open "select * from turnOverDis where (fyear='" & next_Yrs & "' and Current_Next='next')", CCON, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
    
    If RS!NotCreated = "y" Then
    
      If IsDate(txtFromSale_nyr.text) Then
         RS!fromdate = txtFromSale_nyr.text
      Else
         RS!fromdate = Null
      End If
      
      If IsDate(txtToSale_nyr.text) Then
         RS!todate = txtToSale_nyr.text
      Else
         RS!todate = Null
         MsgBox "Invalid Date in (Sale From Date)", vbCritical
         Exit Sub
      End If
      
       RS!fromDateSRet = txtFromSaleRet_nyr.text
       
       If IsDate(txtToSaleRet_nyr.text) Then
          RS!toDateSRet = txtToSaleRet_nyr.text
       Else
          MsgBox "Invalid Date in (Sale Return To Date)", vbCritical
          Exit Sub
       End If
       RS.update
       
    End If
    
    End If


End If


MsgBox "Data Saved...", vbInformation


    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()



If MsgBox("Want to Update ?", vbQuestion + vbYesNo) = vbYes Then
   Screen.MousePointer = vbHourglass
   con.Execute "exec feeItemUpdate"
   Screen.MousePointer = vbDefault
End If


'''===================================================
'con.Execute "delete from INVOICEB_free"
'
'If RS.State = 1 Then RS.close
'RS.Open "select INVOICENO,invoiceDate,BOOKCODE,QUANTITY,agentname,godown,Genledger,SUBLEDGER  from invoiceBQry order by INVOICENO", con
'While RS.EOF = False
'If rs1.State = 1 Then rs1.close
'rs1.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,Apply from KitQry where kitcode='" & RS!Bookcode & "'", con
'While rs1.EOF = False
'
'        con.Execute "insert into INVOICEB_Free(INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
'       " values('" & RS!invoiceNo & "','" & Format(RS!INVOICEDATE, "MM/dd/yyyy") & "','" & Trim(RS!Genledger) & "','" & Trim(RS!SUBLEDGER) & "','" & rs1!Bookcode & "','" & (rs1!qty * RS!QUANTITY) & "','" & rs1!rate & "','" & RS!agentname & "','" & setupid & "','" & session & "','" & RS!Godown & "')"
'
'   rs1.MoveNext
'Wend
'
'RS.MoveNext
'Wend
'
'
'
'
'con.Execute "delete from CREDITB_Free"
'
'
'If RS.State = 1 Then RS.close
'RS.Open "select INVOICENO,invoiceDate,BOOKCODE,QUANTITY,agentname,godown,Genledger,SUBLEDGER from CreditbQry order by INVOICENO", con
'While RS.EOF = False
'If rs1.State = 1 Then rs1.close
'rs1.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,Apply from KitQry where kitcode='" & RS!Bookcode & "'", con
'While rs1.EOF = False
'
'      con.Execute "insert into CREDITB_Free(INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
'       " values('" & RS!invoiceNo & "','" & Format(RS!INVOICEDATE, "MM/dd/yyyy") & "','" & Trim(RS!Genledger) & "','" & Trim(RS!SUBLEDGER) & "','" & rs1!Bookcode & "','" & (rs1!qty * RS!QUANTITY) & "','" & rs1!rate & "','" & RS!agentname & "','" & setupid & "','" & session & "','" & RS!Godown & "')"
'
'
'rs1.MoveNext
'
'Wend
'RS.MoveNext
'Wend
'
'
'con.Execute "delete from INVOICEBSP_Free"
'
'str1 = "select INVOICENO,invoiceDate,BOOKCODE,QUANTITY,agentname,godown,Genledger,SUBLEDGER from invoiceSPBQry order by INVOICENO"
'
'If RS.State = 1 Then RS.close
'RS.Open str1, con
'While RS.EOF = False
'If rs1.State = 1 Then rs1.close
'rs1.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,Apply from KitQry where kitcode='" & RS!Bookcode & "'", con
'While rs1.EOF = False
'
'        con.Execute "insert into INVOICEBSP_Free(INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
'       " values('" & RS!invoiceNo & "','" & Format(RS!INVOICEDATE, "MM/dd/yyyy") & "','" & Trim(RS!Genledger) & "','" & Trim(RS!SUBLEDGER) & "','" & rs1!Bookcode & "','" & (rs1!qty * RS!QUANTITY) & "','" & rs1!rate & "','" & RS!agentname & "','" & setupid & "','" & session & "','" & RS!Godown & "')"
'
'   rs1.MoveNext
'Wend
'RS.MoveNext
'Wend
'
'
'
'con.Execute "delete from INVOICEB_spRet_Free"
'
'str1 = "select INVOICENO,invoiceDate,BOOKCODE,QUANTITY,agentname,godown,Genledger,SUBLEDGER from invoiceSPRETBQry order by INVOICENO"
'
'If RS.State = 1 Then RS.close
'RS.Open str1, con
'
'While RS.EOF = False
'If rs1.State = 1 Then rs1.close
'rs1.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,Apply from KitQry where kitcode='" & RS!Bookcode & "'", con
'While rs1.EOF = False
'
'       con.Execute "insert into INVOICEB_spRet_Free(INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
'       " values('" & RS!invoiceNo & "','" & Format(RS!INVOICEDATE, "MM/dd/yyyy") & "','" & Trim(RS!Genledger) & "','" & Trim(RS!SUBLEDGER) & "','" & rs1!Bookcode & "','" & (rs1!qty * RS!QUANTITY) & "','" & rs1!rate & "','" & RS!agentname & "','" & setupid & "','" & session & "','" & RS!Godown & "')"
'
'rs1.MoveNext
'Wend
'RS.MoveNext
'Wend
'
'
''BookStock_free
'
'con.Execute "delete from BookStock_free"
'
'If RS.State = 1 Then RS.close
'RS.Open "select EntryNo,BOOKCODE,Qty,Dates,Godown_Out from BookStock order by EntryNo", con
'While RS.EOF = False
'If rs1.State = 1 Then rs1.close
'rs1.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,Apply from KitQry where kitcode='" & RS!Bookcode & "'", con
'While rs1.EOF = False
'
'
'    con.Execute "insert into BookStock_free(EntryNo,Dates,BOOKCODE,Qty,setupid,Fyear,Godown) " & _
'    " values('" & RS!EntryNo & "','" & Format(RS.Fields("Dates").value, "MM/dd/yyyy") & "','" & rs1!Bookcode & "','" & Val(RS!qty) & "','" & setupid & "','" & session & "','" & RS!Godown_Out & "')"
'
'
'   rs1.MoveNext
'Wend
'RS.MoveNext
'Wend
'
'
'


End Sub
Private Sub Command4_Click()

''''''Dim CON_next As New ADODB.Connection
''''''Set CON_next = New ADODB.Connection
''''''Dim rs_cr As New ADODB.Recordset
''''''
''''''
''''''If rs1.State = 1 Then rs1.close
''''''rs1.Open "select distinct EntryNo,dates from TurnOver", con
''''''While rs1.EOF = False
''''''   con.Execute "update INVOICEA set toddate='" & Format(rs1!Dates, "MM/dd/yyyy") & "' where todid='" & rs1!EntryNo & "'"
''''''   con.Execute "update creditA set toddate='" & Format(rs1!Dates, "MM/dd/yyyy") & "' where todid='" & rs1!EntryNo & "'"
''''''   con.Execute "update CNF1A set toddate='" & Format(rs1!Dates, "MM/dd/yyyy") & "' where todid='" & rs1!EntryNo & "'"
''''''rs1.MoveNext
''''''Wend
''''''
''''''
''''''If rs1.State = 1 Then rs1.close
''''''rs1.Open "select fromDate,toDate,fyear,Current_Next,DataBase,fromDateSRet,toDateSRet from turnOverDis where Current_Next='next'", CCON
''''''If rs1.EOF = False Then
''''''
''''''
''''''Set CON_next = New ADODB.Connection
''''''If LCase(server_) = "server" Then
''''''   CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID= " & sql_user  & "; PWD=dinesh.123;"
''''''   CON_next.Open
''''''Else
''''''   CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs1!Database & "; UID=; PWD=;"
''''''   CON_next.Open
''''''End If
''''''
''''''End If
''''''
''''''If rs1.State = 1 Then rs1.close
''''''rs1.Open "select distinct EntryNo,dates from TurnOver", con
''''''While rs1.EOF = False
''''''   con.Execute "update INVOICEA set toddate='" & Format(rs1!Dates, "MM/dd/yyyy") & "' where todid='" & rs1!EntryNo & "'"
''''''   CON_next.Execute "update creditA set toddate='" & Format(rs1!Dates, "MM/dd/yyyy") & "' where todid='" & rs1!EntryNo & "'"
''''''   CON_next.Execute "update CNF1A set toddate='" & Format(rs1!Dates, "MM/dd/yyyy") & "' where todid='" & rs1!EntryNo & "'"
''''''rs1.MoveNext
''''''Wend


''''If rs1.State = 1 Then rs1.close
''''rs1.Open "select cnn from CNF1A order by cnn", con, adOpenDynamic, adLockOptimistic
''''While rs1.EOF = False
''''If RS.State = 1 Then RS.close
''''RS.Open "select * from CreditNotDet where cnn=" & rs1!cnn & "", con, adOpenDynamic, adLockOptimistic
''''s10 = ""
''''For I = 1 To RS.RecordCount
''''       If s10 = "" Then
''''           s10 = RS!NARR
''''       Else
''''           s10 = s10 & "," & RS!NARR
''''       End If
''''      RS.MoveNext
''''Next
''''
''''If s10 <> "" Then
''''   con.Execute "update CNF1A set desc_='" & s10 & "' where cnn=" & rs1!cnn & ""
''''End If
''''
''''rs1.MoveNext
''''Wend



If rs1.State = 1 Then rs1.close
rs1.Open "select dnn from DNFA order by dnn", con, adOpenDynamic, adLockOptimistic
While rs1.EOF = False
If RS.State = 1 Then RS.close
RS.Open "select * from DebitNotDet where dnn=" & rs1!dnn & "", con, adOpenDynamic, adLockOptimistic
s10 = ""
For I = 1 To RS.RecordCount
       If s10 = "" Then
           s10 = RS!NARR
       Else
           s10 = s10 & "," & RS!NARR
       End If
      RS.MoveNext
Next
   
If s10 <> "" Then
   con.Execute "update DNFA set desc_='" & Mid(s10, 1, 300) & "' where dnn=" & rs1!dnn & ""
End If
    
rs1.MoveNext
Wend
    
    
End Sub
Private Sub Command5_Click()

Dim headName, headMail As String
Dim sss As New ADODB.Recordset



  
'=====================================

  financialyear
  
  dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & financialyear_Fdate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & financialyear_Tdate & "',103))"
  
  con.Execute "update invoicea set App_Add='n',Appno=''"
  con.Execute "update invoiceb set App_Add='n',Appno=''"
  
  If RS.State = 1 Then RS.close
  RS.Open "select  INVOICENO,Fyear,appno from ApprovalDet where fyear='2020-21' group by INVOICENO,Fyear,appno", con
  While RS.EOF = False
    con.Execute "update AppForm set Fyear='" & RS!fyear & "' where appno=" & RS!appno & ""
    con.Execute "update invoicea set App_Add='y',Appno=" & RS!appno & " where invoiceno=" & RS!invoiceNo & ""
    con.Execute "update invoiceb set App_Add='y',Appno=" & RS!appno & " where invoiceno=" & RS!invoiceNo & ""
    RS.MoveNext
  Wend
  
  '===========================================================
    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & "chitraData_2122" & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
       PopUpValue6 = ""
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & "chitraData_1920" & "; UID=; PWD=;"
       CON_next.Open
    End If
    
  ''CON_next.Execute "update invoicea set App_Add='n',Appno=''"
  ''CON_next.Execute "update invoiceb set App_Add='n',Appno=''"
  If RS.State = 1 Then RS.close
  RS.Open "select  INVOICENO,Fyear,appno from ApprovalDet where fyear='2021-22' group by INVOICENO,Fyear,appno", con
  While RS.EOF = False
    con.Execute "update AppForm set Fyear='" & RS!fyear & "' where appno=" & RS!appno & ""
    CON_next.Execute "update invoicea set App_Add='y',Appno=" & RS!appno & " where invoiceno=" & RS!invoiceNo & ""
    CON_next.Execute "update invoiceb set App_Add='y',Appno=" & RS!appno & " where invoiceno=" & RS!invoiceNo & ""
    RS.MoveNext
  Wend
    
  
  MsgBox "ok"
  
  
End Sub

Private Sub Command6_Click()

Dim rs_ As New ADODB.Recordset
Dim kk_ As Integer

kk_ = 0


If rs1.State = 1 Then rs1.close
rs1.Open "select VoucherDate,VoucherType from tmpVoucher where VoucherType='R' order by VoucherDate"
While rs1.EOF = False
   
   kk_ = 0
   
   If rs_.State = 1 Then rs_.close
   rs_.Open "Select vouchernumber from tmpVoucher where Vouchertype='" + rs1!VoucherType + "' " & _
   " and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & rs1!VoucherDate & "',103) order by vouchernumber", con, adOpenDynamic, adLockReadOnly, adCmdText
   While rs_.EOF = False
   
   kk_ = kk_ + 1
   
   con.Execute "update tmpVoucher set vouchernumber1=" & kk_ & " where vouchernumber= " & rs_!VoucherNumber & " and VoucherType='" & rs1!VoucherType & "' " & _
   " and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & rs1!VoucherDate & "',103)"
   
   rs_.MoveNext
   
   Wend

rs1.MoveNext
Wend







End Sub
Private Sub Command7_Click()

Screen.MousePointer = vbHourglass

vs.Clear
vs.Visible = True
vs.FormatString = "VNo|Narr"
vs.ColWidth(0) = 1200
vs.ColWidth(1) = 2200

Dim s1 As Double
Dim s2 As Double
Dim m1 As Double

Dim arr1, arr2
Dim k1 As Integer

arr1 = Array("invoicea", "casha", "credita")
arr2 = Array("invoicec", "cashc", "creditc")

k1 = 1


For I = 0 To UBound(arr1)

Set rs1 = New ADODB.Recordset
rs1.Open "select invoiceno,[NETAMOUNT],[GAMOUNT],[NETAMOUNT]-[GAMOUNT] as Bal from " + arr1(I) + "  where (convert(smalldatetime,InvoiceDate,103) >= convert(smalldatetime,'" + Trim(txtFromSale_cyr.text) + "',103) and convert(smalldatetime,InvoiceDate,103) <= convert(smalldatetime,'" + Trim(txtToSale_cyr.text) + "',103))", con

While rs1.EOF = False

s1 = 0
s2 = 0

m1 = Round(rs1(3), 2)

If RS.State = 1 Then RS.close
RS.Open "select -1*sum(AMOUNT) as amt from " + arr2(I) + "  where DEBITORCREDIT='Debit' and amount>0 and invoiceno=" & rs1(0), con
If Not IsNull(RS(0)) Then
   s1 = RS(0)
End If

If RS.State = 1 Then RS.close
RS.Open "select sum(AMOUNT) as amt from " + arr2(I) + "  where DEBITORCREDIT='Credit' and amount>0 and invoiceno=" & rs1(0), con
If Not IsNull(RS(0)) Then
   s2 = RS(0)
End If

a1 = Round(m1, 2)
a2 = Round((s1 + s2), 2)

If a1 = a2 Then
Else
 
 vs.rows = vs.rows + 1
 vs.TextMatrix(k1, 0) = rs1!invoiceNo
 vs.TextMatrix(k1, 1) = "" & Round(m1 - (s1 + s2), 2) & " : " & arr1(I)
 DoEvents
 k1 = k1 + 1

End If

rs1.MoveNext
Wend

Next

Dim dr, CR As Double


If RS.State = 1 Then RS.close
RS.Open "SELECT distinct [VoucherType],[VoucherDate],[VoucherNumber] FROM VOUCHERS where (convert(smalldatetime,VoucherDate,103) >= convert(smalldatetime,'" + Trim(txtFromSale_cyr.text) + "',103) and convert(smalldatetime,VoucherDate,103) <= convert(smalldatetime,'" + Trim(txtToSale_cyr.text) + "',103)) order by VoucherType,VoucherNumber,VoucherDate", con
While RS.EOF = False

dr = 0
CR = 0


If rs1.State = 1 Then rs1.close
rs1.Open "select sum(amount) from vouchers where DebitorCredit='D' and vouchertype='" + Trim(RS(0)) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & RS(1) & "',103) And vouchernumber = " + Trim(RS(2))
If Not IsNull(rs1(0)) Then
 dr = rs1(0)
End If

If rs1.State = 1 Then rs1.close
rs1.Open "select sum(amount) from vouchers where DebitorCredit='C' and vouchertype='" + Trim(RS(0)) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & RS(1) & "',103) And vouchernumber = " + Trim(RS(2))
If Not IsNull(rs1(0)) Then
 CR = rs1(0)
End If

If (dr - CR) <> 0 Then

vs.rows = vs.rows + 1
vs.TextMatrix(k1, 0) = rs1!VoucherType & ":" & rs1!VoucherDate & ":" & rs1!VoucherNumber
vs.TextMatrix(k1, 1) = "" & Round((dr - CR), 2) & " : VOUCHERS"

DoEvents
k1 = k1 + 1

End If


RS.MoveNext
Wend


Screen.MousePointer = vbDefault
MsgBox "complete...."

End Sub

Private Sub Form_Load()
 BackColorFrom Me

 
Set RS = New ADODB.Recordset
RS.Open "select toDate,fyear from turnOverDis where Current_Next='next' order by fyear", CCON
If RS.EOF = False Then
   'RS.MoveNext
   next_Yrs = RS!fyear
End If

If LCase(UserName = "admin") Then
   cmdtmpClear.Visible = True
Else
   cmdtmpClear.Visible = False
End If

Command3.Visible = False
If LCase(LCase(UserName) = "v") Then
Command3.Visible = True
Command7.Visible = True
End If
 
 
 'turnOverDis
 Set RS = New ADODB.Recordset
 RS.Open "select * from turnOverDis where fyear='" & session & "'", CCON
 If RS.EOF = False Then
    txtFromSale_cyr = IIf(IsNull(RS!fromdate), "__/__/____", RS!fromdate)
    txtToSale_cyr = IIf(IsNull(RS!todate), "__/__/____", RS!todate)
    
    txtFromSaleRet_cyr = IIf(IsNull(RS!fromDateSRet), "__/__/____", RS!fromDateSRet)
    txtToSaleRet_cyr = IIf(IsNull(RS!toDateSRet), "__/__/____", RS!toDateSRet)
 End If
 
 Set RS = New ADODB.Recordset
 RS.Open "select * from turnOverDis where Current_Next='next'", CCON
 If RS.EOF = False Then
 If RS!NotCreated = "y" Then
    txtFromSale_nyr = IIf(IsNull(RS!fromdate), "__/__/____", RS!fromdate)
    txtToSale_nyr = IIf(IsNull(RS!todate), "__/__/____", RS!todate)
    txtFromSaleRet_nyr = IIf(IsNull(RS!fromDateSRet), "__/__/____", RS!fromDateSRet)
    txtToSaleRet_nyr = IIf(IsNull(RS!toDateSRet), "__/__/____", RS!toDateSRet)
 End If
 End If
 
 
cmdtmpClear.Visible = False
 
End Sub

Private Sub SSTab1_GotFocus()
On Error GoTo errtxt
If SSTab1.Tabs(1).Key = "Tab1" Then
   Frame3.Visible = False
   Frame2.Visible = True
   Set RS = New ADODB.Recordset
   RS.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockOptimistic, adCmdText
   Me.txtCname.text = RS!cname
   Me.add1.text = RS!add1
   Me.add2.text = RS!add2
   Me.city.text = RS!city & ""
   Me.txtPhone1.text = RS!phone1 & ""
   Me.txtphone2.text = RS!phone2 & ""
   Me.txtFax = Trim(RS!FAX) & ""
   txtyarfrom.text = Trim(RS!yarfrom) & ""
   txtyarto.text = Trim(RS!yarto) & ""
   txtEmail.text = IIf(IsNull(Trim(RS!email)), "", Trim(RS!email))
   txtcourt.text = RS!COURT & ""
   txtrem1.text = RS!rem1 & ""
   txtrem2.text = RS!rem2 & ""
   txtbankadvice.text = RS!bankadviceno & ""
   txtuptt.text = RS!uptt & ""
   txtcst.text = RS!cst & ""
   RS.update
End If
errtxt:
If err.Number <> 0 Then
MsgBox "Setup is Invalid. Please Contact to the Vendor."
End If
    
End Sub

