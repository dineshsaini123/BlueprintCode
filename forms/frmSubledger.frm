VERSION 5.00
Begin VB.Form frmSubledger 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Sub Ledger"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10428
   Icon            =   "frmSubledger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   10428
   Begin VB.Frame panel 
      BackColor       =   &H00E0E0E0&
      Height          =   9195
      Left            =   0
      TabIndex        =   33
      Top             =   480
      Width           =   10275
      Begin VB.ComboBox cbomsme 
         Height          =   288
         ItemData        =   "frmSubledger.frx":000C
         Left            =   8460
         List            =   "frmSubledger.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   3492
         Width           =   996
      End
      Begin VB.ComboBox cboPostage 
         Height          =   288
         ItemData        =   "frmSubledger.frx":0023
         Left            =   8460
         List            =   "frmSubledger.frx":002D
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   2736
         Width           =   990
      End
      Begin VB.CommandButton cmdAddSeries 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add S&eries Wise Discount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   3924
         Width           =   1755
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   7335
         Width           =   495
      End
      Begin VB.ComboBox cbofrt 
         Height          =   288
         ItemData        =   "frmSubledger.frx":003A
         Left            =   8448
         List            =   "frmSubledger.frx":0047
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   3105
         Width           =   990
      End
      Begin VB.ComboBox cboProfile 
         Height          =   288
         ItemData        =   "frmSubledger.frx":005B
         Left            =   5625
         List            =   "frmSubledger.frx":005D
         TabIndex        =   21
         Top             =   5445
         Width           =   2040
      End
      Begin VB.TextBox txtStCode 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   5835
         MaxLength       =   25
         TabIndex        =   12
         Top             =   3420
         Width           =   375
      End
      Begin VB.TextBox Phone 
         Height          =   285
         Left            =   3360
         TabIndex        =   14
         Top             =   3795
         Width           =   4305
      End
      Begin VB.TextBox txtpin 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6555
         MaxLength       =   25
         TabIndex        =   13
         Top             =   3435
         Width           =   1095
      End
      Begin VB.TextBox txtgst 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   5940
         MaxLength       =   25
         TabIndex        =   18
         Top             =   4800
         Width           =   1695
      End
      Begin VB.CheckBox Check1_School 
         BackColor       =   &H00C0FFFF&
         Caption         =   "School"
         Height          =   285
         Left            =   7800
         TabIndex        =   69
         Top             =   1260
         Width           =   1590
      End
      Begin VB.CommandButton cmdedit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Replace Rep. Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   5805
         Width           =   1755
      End
      Begin VB.ComboBox cborep6 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3360
         TabIndex        =   27
         Top             =   7380
         Width           =   4335
      End
      Begin VB.ComboBox cborep5 
         Height          =   315
         Left            =   3360
         TabIndex        =   26
         Top             =   7035
         Width           =   4335
      End
      Begin VB.ComboBox cborep4 
         Height          =   315
         Left            =   3360
         TabIndex        =   25
         Top             =   6735
         Width           =   4335
      End
      Begin VB.ComboBox cborep3 
         Height          =   315
         Left            =   3360
         TabIndex        =   24
         Top             =   6435
         Width           =   4335
      End
      Begin VB.ComboBox cborep2 
         Height          =   315
         Left            =   3360
         TabIndex        =   23
         Top             =   6135
         Width           =   4335
      End
      Begin VB.ComboBox cborep1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3360
         TabIndex        =   22
         Top             =   5835
         Width           =   4335
      End
      Begin VB.TextBox txtLessP 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   20
         Top             =   5460
         Width           =   855
      End
      Begin VB.ComboBox cboTrans 
         Height          =   315
         Left            =   3360
         TabIndex        =   19
         Top             =   5100
         Width           =   4335
      End
      Begin VB.TextBox txtPan 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   17
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox txtCityId 
         Height          =   315
         Left            =   7680
         TabIndex        =   57
         Top             =   1980
         Width           =   795
      End
      Begin VB.CheckBox Check1_printer 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Printer/Binder"
         Height          =   285
         Left            =   7800
         TabIndex        =   54
         Top             =   945
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   630
         TabIndex        =   51
         Top             =   7845
         Width           =   7320
         Begin VB.CommandButton cmdAdd_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   30
            Picture         =   "frmSubledger.frx":005F
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   45
            Width           =   1110
         End
         Begin VB.CommandButton cmdSave_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   1155
            Picture         =   "frmSubledger.frx":0C43
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   45
            Width           =   1170
         End
         Begin VB.CommandButton cmdDelete_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   2385
            Picture         =   "frmSubledger.frx":1827
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   45
            Width           =   1155
         End
         Begin VB.CommandButton cmdEdit_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   3600
            Picture         =   "frmSubledger.frx":240B
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   45
            Width           =   1155
         End
         Begin VB.CommandButton cmdExit_12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   5970
            Picture         =   "frmSubledger.frx":2818
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   45
            Width           =   1230
         End
         Begin VB.CommandButton cmdSearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sea&rch"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   8010
            Picture         =   "frmSubledger.frx":33FC
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   45
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   4800
            Picture         =   "frmSubledger.frx":3FE0
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   45
            Width           =   1125
         End
      End
      Begin VB.TextBox txtFindSL 
         Height          =   285
         Left            =   7800
         MaxLength       =   50
         TabIndex        =   50
         Top             =   585
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.ComboBox Comboslgenledgerdiscription 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3360
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   4290
      End
      Begin VB.ComboBox Combosldiscountcategory 
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Top             =   2730
         Width           =   885
      End
      Begin VB.ComboBox Combosldistrictcode 
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3075
         Width           =   4320
      End
      Begin VB.TextBox Textslsubledgerdiscription 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3360
         MaxLength       =   45
         TabIndex        =   1
         Top             =   570
         Width           =   4350
      End
      Begin VB.TextBox Textsldiscriptionforinvoice 
         Height          =   345
         Left            =   3360
         MaxLength       =   39
         TabIndex        =   2
         Top             =   870
         Width           =   4305
      End
      Begin VB.TextBox Textsladdress1 
         Height          =   345
         Left            =   3360
         MaxLength       =   49
         TabIndex        =   3
         Top             =   1230
         Width           =   4305
      End
      Begin VB.TextBox Textslyearopeningbalance 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   345
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2355
         Width           =   4305
      End
      Begin VB.TextBox Textsladdress2 
         Height          =   345
         Left            =   3360
         MaxLength       =   49
         TabIndex        =   4
         Top             =   1590
         Width           =   4305
      End
      Begin VB.TextBox Textsladdress3 
         Height          =   345
         Left            =   3360
         MaxLength       =   49
         TabIndex        =   5
         Top             =   1950
         Width           =   4305
      End
      Begin VB.TextBox txtowner 
         Height          =   285
         Left            =   10425
         TabIndex        =   34
         Top             =   4260
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox txtoffdays 
         Height          =   285
         Left            =   3360
         TabIndex        =   15
         Top             =   4170
         Width           =   4290
      End
      Begin VB.ComboBox cbodisii 
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   2700
         Width           =   945
      End
      Begin VB.ComboBox cbodisii1 
         Height          =   315
         Left            =   5580
         TabIndex        =   9
         Top             =   2700
         Width           =   945
      End
      Begin VB.ComboBox cboState 
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3435
         Width           =   2040
      End
      Begin VB.TextBox txtEmail 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3360
         MaxLength       =   100
         TabIndex        =   16
         Top             =   4500
         Width           =   4275
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "MSME :"
         Height          =   288
         Left            =   7812
         TabIndex        =   80
         Top             =   3540
         Width           =   732
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Postage"
         Height          =   288
         Left            =   7776
         TabIndex        =   79
         Top             =   2772
         Width           =   600
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight"
         Height          =   285
         Left            =   7785
         TabIndex        =   76
         Top             =   3150
         Width           =   600
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Profile"
         Height          =   255
         Index           =   9
         Left            =   4770
         TabIndex        =   73
         Top             =   5460
         Width           =   765
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   255
         Left            =   5400
         TabIndex        =   71
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "GST "
         Height          =   255
         Left            =   5580
         TabIndex        =   70
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Manager Name"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   67
         Top             =   7425
         Width           =   2655
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Rep. Name  5"
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   66
         Top             =   7080
         Width           =   2655
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Rep. Name  4"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   65
         Top             =   6780
         Width           =   2655
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Rep. Name  3"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   64
         Top             =   6480
         Width           =   2655
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Rep. Name  2"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   63
         Top             =   6180
         Width           =   2655
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "State Head"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   62
         Top             =   5895
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage for approcess Amt. to Collected"
         Height          =   375
         Left            =   600
         TabIndex        =   61
         Top             =   5475
         Width           =   2775
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Transport"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   60
         Top             =   5205
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PIN "
         Height          =   255
         Left            =   6240
         TabIndex        =   59
         Top             =   3435
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PAN"
         Height          =   255
         Left            =   600
         TabIndex        =   58
         Top             =   4875
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "City            (F2 For Search City ....)"
         Height          =   375
         Left            =   600
         TabIndex        =   56
         Top             =   1980
         Width           =   2640
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "* Required fields"
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
         Left            =   240
         TabIndex        =   53
         Top             =   8880
         Width           =   2955
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   1050
         Left            =   585
         Top             =   7800
         Width           =   7470
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "District Name "
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   48
         Top             =   3180
         Width           =   2655
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Category                         (I)"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   47
         Top             =   2820
         Width           =   3015
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "(Description for Invoice)"
         Height          =   375
         Left            =   600
         TabIndex        =   46
         Top             =   930
         Width           =   2295
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "* Sub. Ledger Discription"
         Height          =   375
         Left            =   600
         TabIndex        =   45
         Top             =   570
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   375
         Left            =   630
         TabIndex        =   44
         Top             =   1260
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "* General Ledger Description"
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   180
         Width           =   2595
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Year Opening Balance"
         Height          =   255
         Left            =   630
         TabIndex        =   42
         Top             =   2460
         Width           =   2505
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone "
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   3870
         Width           =   2295
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "E- Mail"
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   4515
         Width           =   1335
      End
      Begin VB.Label Label36 
         Caption         =   "Owner"
         Height          =   255
         Left            =   10380
         TabIndex        =   39
         Top             =   3780
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   " (II)"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   38
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   " (III)"
         Height          =   255
         Index           =   2
         Left            =   5340
         TabIndex        =   37
         Top             =   2760
         Width           =   315
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   255
         Index           =   1
         Left            =   585
         TabIndex        =   36
         Top             =   3555
         Width           =   2655
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile "
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   4200
         Width           =   3015
      End
   End
   Begin VB.Label header 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   19185
   End
End
Attribute VB_Name = "frmSubledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit_ As Boolean
Dim str1 As String

Private Sub cbodisii1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'   cboState.SetFocus
End If
End Sub
Private Sub cmdAdd_1_Click()

str1 = "Save"
clearFrom frmSubledger
Comboslgenledgerdiscription.SetFocus
Check1_school.value = 0
cbofrt.ListIndex = 0
cboPostage.ListIndex = 0

cbomsme.ListIndex = 0

ButtonPermissionNew cmdSave_2, cmdDelete_3, cmdEdit_4, "menusubleadgermaster"

End Sub

Private Sub cmdAddSeries_Click()

If LCase(UserName) <> "admin" Then
   MsgBox "You can'nt open this form !!", vbExclamation, "Alert"
   Exit Sub
End If


If Textslsubledgerdiscription.text <> "" Then
    PopUpValue6 = Textslsubledgerdiscription.text
    frmSeriesWiseDis.Show 1
Else
   MsgBox "Plz Search Party ..", vbInformation
End If

End Sub
Private Sub cmdDelete_3_Click()

If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then
    
    con.BeginTrans
    con.Execute "delete from  SLEDGER where SUBLEDGER='" & txtFindSL & "' and " & stringyear
    CCON.Execute "delete from  SLEDGER where SUBLEDGER='" & txtFindSL & "'"
    con.CommitTrans
    
    
    
    
    
    createLog UserName, "" & Mid(Textslsubledgerdiscription.text, 1, 5), "SLEDGER ", " Delete : OP : " & Textslyearopeningbalance.text, Date
   
    Call cmdAdd_1_Click
    
End If


End Sub

Private Sub cmdEdit_4_Click()
    edit_ = True
    str1 = "Modify"
    cmdEdit_4.Enabled = False
    cmdDelete_3.Enabled = False
    cmdSave_2.Enabled = True
    Textslsubledgerdiscription.SetFocus
    
    If (LCase(UserName) = "y" Or LCase(UserName) = "v" Or LCase(UserName) = "admin") Then
        cmdDelete_3.Enabled = True
    End If
    
    
End Sub
Private Sub cmdedit_Click()
frmReplaceName.Show 1
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdSave_2_Click()


 If edit_ = False Then str1 = "Save"


  If Comboslgenledgerdiscription.text = "" Then
      MsgBox "Select Gen. Ledger ...", vbCritical
      Comboslgenledgerdiscription.SetFocus
      Exit Sub
  End If

  If Textslsubledgerdiscription.text = "" Then
      MsgBox "Enter Subledger Ledger ...", vbCritical
      Textslsubledgerdiscription.SetFocus
      Exit Sub
  End If
  
  
 If Comboslgenledgerdiscription = "SUNDRY DEBTORS" Then
    If Textsladdress3 = "" Then
          MsgBox "Select City Name", vbCritical
          Textsladdress3.SetFocus
          Exit Sub
    End If
 End If


 If MsgBox("want to " & str1 & " ?", vbQuestion + vbYesNo) = vbYes Then
    
  If str1 = "Save" Then
    
    If RS.State = 1 Then RS.close
    RS.Open "select code from SLEDGER where code='" & Trim(Mid(Textslsubledgerdiscription.text, 1, 6)) & "'", con
    If RS.EOF = False Then
       MsgBox "This Code Already Exist...", vbCritical
       Textslsubledgerdiscription.SetFocus
       Exit Sub
    End If
    saveData
  Else
    ModifyData
  End If
  
  
  
  On Error Resume Next
  con.Execute "exec RepChange 1,'" & txtFindSL & "','" & cborep1 & "','" & cborep2 & "','" & cborep3 & "','" & cborep4 & "','" & cborep5 & "','" & cborep6 & "' "
  con.Execute "exec RepChange 2,'" & txtFindSL & "','" & cborep1 & "','" & cborep2 & "','" & cborep3 & "','" & cborep4 & "','" & cborep5 & "','" & cborep6 & "' "
  con.Execute "exec RepChange 3,'" & txtFindSL & "','" & cborep1 & "','" & cborep2 & "','" & cborep3 & "','" & cborep4 & "','" & cborep5 & "','" & cborep6 & "' "
  
  
  
  
  cmdDelete_3.Enabled = False
  cmdSave_2.Enabled = False
 End If

End Sub
Sub ModifyData()


    If RS.State = 1 Then RS.close
    RS.Open "select * from SLEDGER where " & stringyear & " and SUBLEDGER='" & txtFindSL & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
            
            RS!freight = cbofrt.text
            RS!postage = cboPostage.text
            
            RS!MSME = cbomsme.text
            
            RS!profile_ = cboProfile.text
            If Check1_printer.value = 1 Then
               RS!printer_binder = "y"
            Else
               RS!printer_binder = "n"
            End If
            
            If Check1_school.value = 1 Then
               RS!bookseller_sch = "S"
            Else
               RS!bookseller_sch = "B"
            End If
            
            RS!statecode = Trim(txtStCode.text)
            
            RS!lessp = Trim(txtLessP)
            RS!transport = UCase(cboTrans.text)
            RS!pin = Trim(txtPin)

            RS!states = cbostate.text
            RS!email = txtEmail.text
            RS!gledger = Comboslgenledgerdiscription.text
            RS!subledger = Textslsubledgerdiscription.text
                                   
            RS!party = Trim(Mid(Textslsubledgerdiscription.text, 7))
            RS!Code = Trim(Left(Textslsubledgerdiscription.text, 5))
            RS!DESCFORINVOICE = Textsldiscriptionforinvoice.text
            RS!YEAROPENING = Val(Textslyearopeningbalance.text)
            'If Trim(Textsladdress1.Text) <> "" Then
            RS!address1 = Trim(Textsladdress1.text)
            'End If
            'If Trim(Textsladdress2.Text) <> "" Then
             RS!address2 = Trim(Textsladdress2.text)
            'End If
            'If Trim(Textsladdress3.Text) <> "" Then
             RS!address3 = Trim(Textsladdress3.text)
            'End If
            If Combosldiscountcategory.text <> "" Then
                RS!DISCATEGORY = Combosldiscountcategory.text
            Else
                RS!DISCATEGORY = ""
            End If
            If Combosldistrictcode.text <> "" Then
                RS!distcode = Combosldistrictcode.text
            Else
                RS!distcode = ""
            End If
            
             If phone.text <> "" Then
                RS!phone = phone.text
            Else
                RS!phone = ""
            End If
            
            If txtoffdays.text <> "" Then
                RS!mobile = txtoffdays.text
            Else
                RS!mobile = ""
            End If
            RS!category2 = cbodisii.text
            RS!Category3 = cbodisii1.text
            RS!cityId = txtCityId
            RS!pan = Trim(txtPan)
            RS!gst = Trim(txtgst.text)
            'RS!cityname = Textsladdress3.Text
                
            RS!RepName1 = Trim(cborep1)
            RS!RepName2 = Trim(cborep2)
            RS!RepName3 = Trim(cborep3)
            RS!RepName4 = Trim(cborep4)
            RS!RepName5 = Trim(cborep5)
            RS!RepName6 = Trim(cborep6)

    
                
            RS.update
            
            con.Execute "exec UpdateLedger 'VOUCHERS','" & txtFindSL & "','" & Textslsubledgerdiscription.text & "','" & session & "','" & main.setupid & "'"
            
            con.Execute "update CASHA set SUBLEDGER='" & Textslsubledgerdiscription.text & "' where subledger='" & txtFindSL & "'"
            con.Execute "update CASHB set SUBLEDGER='" & Textslsubledgerdiscription.text & "' where subledger='" & txtFindSL & "'"
            
             
            con.Execute "update ReceiveIssueParty set PartyName='" & Textslsubledgerdiscription.text & "' where PartyName='" & txtFindSL & "'"
            
            con.Execute "update ApprovalDet set SUBLEDGER='" & Textslsubledgerdiscription.text & "' where SUBLEDGER='" & txtFindSL & "'"
            addmaster_addSingleData "sledger", Textslsubledgerdiscription, txtFindSL
            
            
            createLog UserName, "" & Mid(Textslsubledgerdiscription.text, 1, 5), "SLEDGER ", " Modify : OP : " & Textslyearopeningbalance.text, Date
            
        End If

End Sub
Sub saveData()

    If RS.State = 1 Then RS.close
    RS.Open "select * from SLEDGER where " & stringyear & " and SUBLEDGER='" & Textslsubledgerdiscription.text & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
    
            RS.AddNew
            
            RS!postage = cboPostage.text
            RS!freight = cbofrt.text
            
            RS!MSME = cbomsme.text
            
            RS!profile_ = cboProfile.text
            
            If Check1_printer.value = 1 Then
               RS!printer_binder = "y"
            Else
               RS!printer_binder = "n"
            End If
            
            If Check1_school.value = 1 Then
               RS!bookseller_sch = "S"
            Else
               RS!bookseller_sch = "B"
            End If
            
            
            RS!statecode = Trim(txtStCode.text)
            
            RS!lessp = Trim(txtLessP)
            RS!pin = Trim(txtPin)
            
            RS!transport = UCase(cboTrans.text)
            RS!states = cbostate.text
            RS!states = cbostate.text
            RS!email = txtEmail.text
            RS!gledger = Comboslgenledgerdiscription.text
            RS!subledger = Textslsubledgerdiscription.text
            RS!party = Trim(Mid(Textslsubledgerdiscription.text, 7))
            RS!Code = Trim(Left(Textslsubledgerdiscription.text, 5))
            RS!DESCFORINVOICE = Textsldiscriptionforinvoice.text
            RS!YEAROPENING = Val(Textslyearopeningbalance.text)
            If Trim(Textsladdress1.text) <> "" Then
                RS!address1 = Trim(Textsladdress1.text)
            End If
            If Trim(Textsladdress2.text) <> "" Then
                RS!address2 = Trim(Textsladdress2.text)
            End If
            If Trim(Textsladdress3.text) <> "" Then
                RS!address3 = Trim(Textsladdress3.text)
            End If
            If Combosldiscountcategory.text <> "" Then
                RS!DISCATEGORY = Combosldiscountcategory.text
            Else
                RS!DISCATEGORY = ""
            End If
            If Combosldistrictcode.text <> "" Then
                RS!distcode = Combosldistrictcode.text
            Else
                RS!distcode = ""
            End If
            
             If phone.text <> "" Then
                RS!phone = phone.text
            Else
                RS!phone = ""
            End If
            
            If txtoffdays.text <> "" Then
                RS!mobile = txtoffdays.text
            Else
                RS!offdays = ""
            End If
            RS!category2 = cbodisii.text
            RS!Category3 = cbodisii1.text
            
            RS!fyear = session
            RS!setupid = setupid
            
            RS!cityId = txtCityId
            RS!pan = Trim(txtPan)
            RS!gst = Trim(txtgst.text)
            
            RS!RepName1 = Trim(cborep1)
            RS!RepName2 = Trim(cborep2)
            RS!RepName3 = Trim(cborep3)
            RS!RepName4 = Trim(cborep4)
            RS!RepName5 = Trim(cborep5)
            RS!RepName6 = Trim(cborep6)
            
                
            RS.update
            
            addmaster_addSingleData "sledger", Textslsubledgerdiscription, Textslsubledgerdiscription
            
            
            createLog UserName, "" & Mid(Textslsubledgerdiscription.text, 1, 5), "SLEDGER ", " Saved : OP : " & Textslyearopeningbalance.text, Date
        
        End If

End Sub
Sub search_Data()

    If RS.State = 1 Then RS.close
    RS.Open "select * from SLEDGER where " & stringyear & " and SUBLEDGER='" & txtFindSL.text & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
    
            If Not IsNull(RS!freight) Then
            
            If Len(RS!freight) > 0 Then
               cbofrt.text = RS!freight & ""
            End If
            End If
            
            If Not IsNull(RS!postage) Then
            If Len(RS!postage) > 0 Then
               cboPostage.text = RS!postage & ""
            End If
            
            End If
            
            
            If Not IsNull(RS!MSME) Then
            If Len(RS!MSME) > 0 Then
               cbomsme.text = RS!MSME & ""
            End If
            
            End If
            
            
           
            
            
            
            cboProfile.text = RS!profile_ & ""
            
            txtStCode.text = RS!statecode & ""
            
            cboTrans.text = RS!transport & ""
            txtLessP = RS!lessp & ""
            
            Comboslgenledgerdiscription.text = RS!gledger
            Textslsubledgerdiscription.text = RS!subledger
            Textsldiscriptionforinvoice.text = RS!DESCFORINVOICE & ""
            Textslyearopeningbalance.text = RS!YEAROPENING
            Textsladdress1.text = RS!address1 & ""
            Textsladdress2.text = RS!address2 & ""
            Textsladdress3.text = RS!address3 & ""
            Combosldiscountcategory.text = RS!DISCATEGORY & ""
            Combosldistrictcode.text = RS!distcode & ""
            phone.text = RS!phone & ""
            txtoffdays.text = RS!mobile & ""
            
            cbodisii.text = RS!category2 & ""
            cbodisii1.text = RS!Category3 & ""
            cbostate.text = RS!states & ""
            txtEmail.text = RS!email & ""
            
            If RS!printer_binder = "y" Then
               Check1_printer.value = 1
            Else
               Check1_printer.value = 0
            End If
            
            If RS!bookseller_sch = "S" Then
               Check1_school.value = 1
            Else
               Check1_school.value = 0
            End If
            
            
            
            If Not IsNull(RS!cityId) Then
              txtCityId = RS!cityId
            End If
            
            txtPan = RS!pan & ""
            txtPin = RS!pin & ""
            txtgst.text = RS!gst & ""
            
            'If Len(RS!RepName1) > 0 Then
             cborep1.text = RS!RepName1 & ""
            'End If
            
            'If Len(RS!RepName2) > 0 Then
               cborep2.text = RS!RepName2 & ""
            'End If
            
            'If Len(RS!RepName3) > 0 Then
               cborep3.text = RS!RepName3 & ""
            'End If
            
            'If Len(RS!RepName4) > 0 Then
               cborep4.text = RS!RepName4 & ""
            'End If
          
            cborep5.text = RS!RepName5 & ""
            cborep6.text = RS!RepName6 & ""


            
            
        End If

End Sub
Private Sub cmdSearch_Click()
popuplist_client "select Subledger,Gledger,YEAROPENING,DESCFORINVOICE from Sledger where " & stringyear & " order by Subledger", CCON
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = True
cmdSave_2.Enabled = False
End Sub
Private Sub Combosldiscountcategory_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Len(Combosldiscountcategory.text) = 0 Then
   MsgBox "Select Category ...", vbCritical
   Combosldiscountcategory.SetFocus
End If
End If
End Sub

Private Sub Comboslgenledgerdiscription_LostFocus()
 
 Check1_school.Visible = False
 If Comboslgenledgerdiscription.text = "SUNDRY DEBTORS" Then
    Check1_school.Visible = True
 End If
 
End Sub

Private Sub Command11_Click()

HeadTbl = "manager"
frmMasters.Show 1
End Sub

Private Sub Commandsearch_Click()
    searchType = "party"
    popuplist_client "select Subledger,Gledger,DESCFORINVOICE from Sledger where " & stringyear & " order by Subledger", CCON
    cmdDelete_3.Enabled = False
    cmdEdit_4.Enabled = True
    cmdSave_2.Enabled = False
End Sub
Private Sub Commandsearch_GotFocus()

If PopUpValue1 <> "" Then
    
    Textslsubledgerdiscription = PopUpValue1
    txtFindSL.text = PopUpValue1
   
    search_Data
    
    Check1_school.Visible = False
    If Comboslgenledgerdiscription.text = "SUNDRY DEBTORS" Then
    Check1_school.Visible = True
    End If
   
    Textslsubledgerdiscription.SetFocus
    
    ButtonPermissionNew cmdSave_2, cmdDelete_3, cmdEdit_4, "menusubleadgermaster"
   
    
End If
             
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys "{tab}"
End Sub
Private Sub Form_Load()

    Me.top = 100
    Me.Left = 100
    Me.Width = 10845
    Me.Height = 10350
    
    cmdedit.Visible = False
    If LCase(UserName) = "admin" Then
       cmdedit.Visible = True
    End If
    
   

    Set RS = New ADODB.Recordset
    RS.Open "select gledger from gledger where " & stringyear & " and slf=1 group by gledger", con, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
        Do While Not RS.EOF
            Comboslgenledgerdiscription.AddItem RS!gledger
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close

    Set RS = New ADODB.Recordset
    RS.Open "select DISTRICTNAME from DISTRICTS where " & stringyear & "  order by DISTRICTNAME", con, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
        Do While Not RS.EOF
            Combosldistrictcode.AddItem RS!DISTRICTNAME
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close


Set RS = New ADODB.Recordset
    RS.Open "select states from states order by states", con, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
        Do While Not RS.EOF
            cbostate.AddItem RS!states
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    
    
    
     
    RS.Open "select distinct categorycode from DISCCATS order by categorycode", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combosldiscountcategory.AddItem RS!categorycode
            Me.cbodisii.AddItem RS!categorycode
            Me.cbodisii1.AddItem RS!categorycode
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    

 'BackColorFrom Me, 1
 
If RS.State = 1 Then RS.close
RS.Open "select  transportname from transportMaster order by transportname", con, adOpenDynamic, adLockReadOnly, adCmdText
cboTrans.Clear
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cboTrans.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If
RS.close
 
 
RS.Open "select Rep as Representative from SalesRepQry order by Rep", CON_blue
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cborep1.AddItem RS(0)
        Me.cborep2.AddItem RS(0)
        Me.cborep3.AddItem RS(0)
        Me.cborep4.AddItem RS(0)
        Me.cborep5.AddItem RS(0)
        'Me.cborep6.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If
 
'======================================
If RS.State = 1 Then RS.close
RS.Open "select distinct Profile_ from SLEDGER order by Profile_", con
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cboProfile.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If
 
 
If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='manager'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   cborep6.AddItem RS(0)
   RS.MoveNext
Wend


If party_name <> "" Then
   search_
End If
 

End Sub
Sub search_()
    
    Textslsubledgerdiscription = party_name
    txtFindSL.text = party_name
 
    search_Data
    
    Check1_school.Visible = False
    If Comboslgenledgerdiscription.text = "SUNDRY DEBTORS" Then
    Check1_school.Visible = True
    End If
   
   
    
   
   
    ButtonPermissionNew cmdSave_2, cmdDelete_3, cmdEdit_4, "menusubleadgermaster"
     
     
    party_name = ""
End Sub

Private Sub sledger_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2


End Sub

Private Sub Textsladdress1_LostFocus()
Textsladdress1 = UCase(Textsladdress1)
End Sub
Private Sub Textsladdress2_LostFocus()
Textsladdress2 = UCase(Textsladdress2)
End Sub

Private Sub Textsladdress3_GotFocus()
If PopUpValue1 <> "" Then

    Textsladdress3 = PopUpValue1
    Combosldistrictcode = PopUpValue2
    cbostate = PopUpValue3
    txtCityId = popupvalue4
    txtStCode = popupvalue5
    
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""
    popupvalue5 = ""

End If
End Sub

Private Sub Textsladdress3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   popuplist10 "select City,District,[State],cityId,StCode from qryMap_ind order by city", CON_blue
   
   'searchType = "city"
   'popuplistFast "select City,District,[State],cityId,StCode from qryMap_ind order by city", con, , , "city..."

End If


If KeyCode = 13 Then
If Comboslgenledgerdiscription = "SUNDRY DEBTORS" Then
   If Textsladdress3 = "" Then
         MsgBox "Select City Name", vbCritical
         Textsladdress3.SetFocus
         Exit Sub
   End If
End If
End If

End Sub

Private Sub Textsladdress3_LostFocus()
Textsladdress3 = UCase(Textsladdress3)
End Sub

Private Sub Textsldiscriptionforinvoice_LostFocus()
Textsldiscriptionforinvoice = UCase(Textsldiscriptionforinvoice)
End Sub

Private Sub Textslsubledgerdiscription_GotFocus()
 
 If PopUpValue1 <> "" Then
    
    Textslsubledgerdiscription = PopUpValue1
    txtFindSL = PopUpValue1
    search_Data
    
  
        
End If
             
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""
             
End Sub

Private Sub Textslsubledgerdiscription_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then

    popuplist_client "select Subledger,Gledger,YEAROPENING,DESCFORINVOICE from Sledger where " & stringyear & " order by Subledger", CCON
    cmdDelete_3.Enabled = False
    cmdEdit_4.Enabled = True
    cmdSave_2.Enabled = False

End If


End Sub

Private Sub Textslsubledgerdiscription_LostFocus()
Textslsubledgerdiscription = UCase(Textslsubledgerdiscription)
End Sub
Private Sub txtEmail_LostFocus()
  txtEmail = LCase(txtEmail)
End Sub
Private Sub txtgst_LostFocus()
   txtgst = UCase(txtgst)
End Sub

Private Sub txtPan_LostFocus()
   txtPan = UCase(txtPan)
End Sub
