VERSION 5.00
Begin VB.Form frmAcc_subledger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Ledger"
   ClientHeight    =   4080
   ClientLeft      =   5028
   ClientTop       =   2688
   ClientWidth     =   8040
   Icon            =   "frmAcc_subledger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8040
   Begin VB.TextBox txtDescInvoice 
      Height          =   285
      Left            =   2205
      MaxLength       =   40
      TabIndex        =   2
      Top             =   1395
      Width           =   4275
   End
   Begin VB.TextBox txtSubledger 
      Height          =   285
      Left            =   2220
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1020
      Width           =   4275
   End
   Begin VB.TextBox txtAdd1 
      Height          =   285
      Left            =   10860
      MaxLength       =   40
      TabIndex        =   38
      Top             =   6540
      Width           =   2895
   End
   Begin VB.TextBox txtAdd2 
      Height          =   285
      Left            =   10860
      MaxLength       =   40
      TabIndex        =   37
      Top             =   6960
      Width           =   2895
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   10860
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   36
      Top             =   8160
      Width           =   2895
   End
   Begin VB.TextBox txtDistrict 
      Height          =   285
      Left            =   10860
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   35
      Top             =   7800
      Width           =   2895
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   10860
      MaxLength       =   30
      TabIndex        =   34
      Top             =   7380
      Width           =   2175
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   10860
      MaxLength       =   30
      TabIndex        =   33
      Top             =   8580
      Width           =   2895
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   10860
      MaxLength       =   30
      TabIndex        =   32
      Top             =   9060
      Width           =   2895
   End
   Begin VB.TextBox txtOpening 
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1815
      Width           =   2895
   End
   Begin VB.TextBox txtRAdd1 
      Height          =   285
      Left            =   15780
      MaxLength       =   40
      TabIndex        =   31
      Top             =   6600
      Width           =   3375
   End
   Begin VB.TextBox txtRAdd2 
      Height          =   285
      Left            =   15780
      MaxLength       =   40
      TabIndex        =   30
      Top             =   7020
      Width           =   3375
   End
   Begin VB.TextBox txtRState 
      Height          =   285
      Left            =   15780
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   29
      Top             =   7860
      Width           =   3375
   End
   Begin VB.TextBox txtRCity 
      Height          =   285
      Left            =   15780
      MaxLength       =   30
      TabIndex        =   28
      Top             =   7440
      Width           =   2655
   End
   Begin VB.TextBox txtRPhone 
      Height          =   285
      Left            =   15780
      MaxLength       =   30
      TabIndex        =   27
      Top             =   8280
      Width           =   3375
   End
   Begin VB.TextBox txtRFax 
      Height          =   285
      Left            =   15780
      MaxLength       =   30
      TabIndex        =   26
      Top             =   8700
      Width           =   3375
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   25
      Top             =   5520
      Width           =   2895
   End
   Begin VB.TextBox txtTpt 
      Height          =   285
      Left            =   2220
      MaxLength       =   20
      TabIndex        =   24
      Top             =   6000
      Width           =   2895
   End
   Begin VB.TextBox txtRange 
      Height          =   285
      Left            =   2220
      MaxLength       =   20
      TabIndex        =   23
      Top             =   6420
      Width           =   2895
   End
   Begin VB.TextBox txtBank 
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   22
      Top             =   6840
      Width           =   2895
   End
   Begin VB.TextBox txtBAdd1 
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   21
      Top             =   7200
      Width           =   2895
   End
   Begin VB.TextBox txtRemark 
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   20
      Top             =   7620
      Width           =   2895
   End
   Begin VB.TextBox txtContact 
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   19
      Top             =   8040
      Width           =   2895
   End
   Begin VB.TextBox txtBPhone 
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   18
      Top             =   8460
      Width           =   2895
   End
   Begin VB.TextBox txtTin 
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   17
      Top             =   8880
      Width           =   2895
   End
   Begin VB.ComboBox txtgenledger 
      Height          =   315
      ItemData        =   "frmAcc_subledger.frx":000C
      Left            =   2220
      List            =   "frmAcc_subledger.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   2955
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   300
      TabIndex        =   16
      Top             =   2760
      Width           =   7575
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
         Left            =   45
         Picture         =   "frmAcc_subledger.frx":0010
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
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
         Left            =   1290
         Picture         =   "frmAcc_subledger.frx":0BF4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1230
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
         Left            =   2520
         Picture         =   "frmAcc_subledger.frx":17D8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1230
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
         Left            =   3750
         Picture         =   "frmAcc_subledger.frx":23BC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1230
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
         Left            =   6240
         Picture         =   "frmAcc_subledger.frx":27C9
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "S&earch"
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
         Left            =   4995
         Picture         =   "frmAcc_subledger.frx":33AD
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   14040
      TabIndex        =   12
      Top             =   4920
      Width           =   735
      Begin VB.OptionButton Option_ss 
         Caption         =   "Super Stockist"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option_dealer 
         Caption         =   "Dealer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option_retairs 
         Caption         =   "Retailers"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3540
         TabIndex        =   13
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.TextBox txtbcityid 
      Height          =   285
      Left            =   13020
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   7380
      Width           =   735
   End
   Begin VB.TextBox txtRcityid 
      Height          =   285
      Left            =   12660
      TabIndex        =   10
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label lblSubledger 
      Height          =   285
      Left            =   6570
      TabIndex        =   71
      Top             =   990
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "* Desc. Invoice :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   405
      TabIndex        =   70
      Top             =   1395
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Opening Balance :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   420
      TabIndex        =   69
      Top             =   1875
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "* Sub Ledger :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   68
      Top             =   1020
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Address1 : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9060
      TabIndex        =   67
      Top             =   6540
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Address2 : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   9060
      TabIndex        =   66
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Range :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   420
      TabIndex        =   65
      Top             =   6420
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "District : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   9060
      TabIndex        =   64
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "City :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   9060
      TabIndex        =   63
      Top             =   7380
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Phone Nos :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   9060
      TabIndex        =   62
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "State :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   9060
      TabIndex        =   61
      Top             =   8220
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Fax Nos :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   9060
      TabIndex        =   60
      Top             =   9060
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Opening Balance : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   9060
      TabIndex        =   59
      Top             =   9480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Address1 : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   14220
      TabIndex        =   58
      Top             =   6600
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Address2 : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   14220
      TabIndex        =   57
      Top             =   7020
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "State : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   14220
      TabIndex        =   56
      Top             =   7860
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "City :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   14220
      TabIndex        =   55
      Top             =   7440
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Phone Nos :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   14220
      TabIndex        =   54
      Top             =   8280
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Fax Nos :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   16
      Left            =   14220
      TabIndex        =   53
      Top             =   8700
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "E-mail ID :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   17
      Left            =   420
      TabIndex        =   52
      Top             =   5580
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "CST :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   18
      Left            =   420
      TabIndex        =   51
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "* Gen Ledger :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   19
      Left            =   420
      TabIndex        =   50
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Bank Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   20
      Left            =   420
      TabIndex        =   49
      Top             =   6780
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Address1 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   21
      Left            =   420
      TabIndex        =   48
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Remark : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   22
      Left            =   420
      TabIndex        =   47
      Top             =   7620
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Contact Person :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   23
      Left            =   420
      TabIndex        =   46
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "UPTT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   24
      Left            =   420
      TabIndex        =   45
      Top             =   8460
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Tin No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   25
      Left            =   420
      TabIndex        =   44
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Billing Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   9060
      TabIndex        =   43
      Top             =   6000
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Residence Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   14220
      TabIndex        =   42
      Top             =   6000
      Width           =   5955
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   360
      TabIndex        =   41
      Top             =   5280
      Width           =   11355
   End
   Begin VB.Label header 
      BackColor       =   &H8000000D&
      Caption         =   "     Sub Ledger"
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
      TabIndex        =   40
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label5 
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
      Left            =   5520
      TabIndex        =   39
      Top             =   8340
      Width           =   2955
   End
End
Attribute VB_Name = "frmAcc_subledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim editval As Boolean
Private Sub cmdAdd_1_Click()

Dim o As Object
For Each o In Me
  If TypeOf o Is TextBox Then
     o.text = ""
  End If
Next

cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
txtSubledger.Enabled = True
txtgenledger.SetFocus
 
End Sub
Sub setDefaults()
    
    Dim o As Object
    For Each o In Me
          If TypeOf o Is TextBox Then
          If o.text = "" Then
             o.text = ""
          End If
          End If
    Next
End Sub
Private Sub cmdDelete_3_Click()

On Error GoTo save:

If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.BeginTrans
   con.Execute "delete from SLEDGER where " & stringyear & " and SUBLEDGER='" & txtSubledger & "'"
   con.CommitTrans
   Call cmdAdd_1_Click
End If

Exit Sub

save:
con.RollbackTrans
MsgBox "" & err.Description

End Sub
Private Sub cmdEdit_4_Click()
 
editval = True
txtgenledger.SetFocus
'txtSubledger.Enabled = False
cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdSave_2_Click()

Dim TypeOfCust As String

If txtgenledger.text = "" Then
   MsgBox "Plz. Select Gen Ledger ...", vbCritical
   txtgenledger.SetFocus
   Exit Sub
End If

If txtSubledger.text = "" Then
   MsgBox "Plz. Enter Subledger ...", vbCritical
   txtSubledger.SetFocus
   Exit Sub
End If



If Option_ss.value = True Then
   TypeOfCust = Option_ss.Caption
ElseIf Option_dealer.value = True Then
   TypeOfCust = Option_dealer.Caption
ElseIf Option_retairs.value = True Then
   TypeOfCust = Option_retairs.Caption
End If

setDefaults

On Error GoTo save:

If editval = False Then



con.BeginTrans

con.Execute "exec insertData_subledger '" & txtgenledger & "','" & txtSubledger & "','" & txtAdd1 & "'," & _
"'" & txtAdd2 & "','" & txtPhone & "','" & txtFax & "'," & Val(txtOpening) & ",'" & txtbcityid & "'," & _
"'" & txtRAdd1 & "','" & txtRAdd2 & "','" & (txtRcityid) & "','" & txtRPhone & "','" & txtRFax & "'," & _
"'" & txtEmail & "','" & txtTpt & "','" & txtRange & "','" & txtBank & "','" & txtBAdd1 & "'," & _
"'" & txtRemark & "','" & txtContact & "','" & txtBPhone & "','" & txtTin & "','" & TypeOfCust & "','" & main.session & "'," & setupid & ",'" & txtDescInvoice.text & "'"

con.CommitTrans

MsgBox "Data Saved ...", vbInformation
cmdDelete_3.Enabled = True
cmdEdit_4.Enabled = True


Else

con.Execute "exec UpdateData_subledger '" & txtgenledger & "','" & txtSubledger & "','" & txtAdd1 & "'," & _
"'" & txtAdd2 & "','" & txtPhone & "','" & txtFax & "'," & Val(txtOpening) & ",'" & txtbcityid & "'," & _
"'" & txtRAdd1 & "','" & txtRAdd2 & "','" & (txtRcityid) & "','" & txtRPhone & "','" & txtRFax & "'," & _
"'" & txtEmail & "','" & txtTpt & "','" & txtRange & "','" & txtBank & "','" & txtBAdd1 & "'," & _
"'" & txtRemark & "','" & txtContact & "','" & txtBPhone & "','" & txtTin & "','" & TypeOfCust & "','" & main.session & "'," & main.setupid & ",'" & txtDescInvoice.text & "','" & lblSubledger.Caption & "'"

Modifytbl lblSubledger.Caption, txtSubledger

MsgBox "Data Modified", vbInformation

cmdDelete_3.Enabled = True
cmdEdit_4.Enabled = True
editval = False

End If

Call cmdAdd_1_Click

Exit Sub

save:

con.RollbackTrans
If err.Number = "-2147217900" Then
   MsgBox err.Description     ' "Subledger Duplicate ...", vbCritical
   txtSubledger.SetFocus
End If

End Sub

Private Sub cmdSearch_Click()
 'tblNo = 1

 popuplist10 "select Subledger,Gledger,YEAROPENING,DESCFORINVOICE from Sledger where " & stringyear & " order by Subledger", con


cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = True
cmdSave_2.Enabled = False
'frmSearchItem.Show
End Sub

Private Sub cmdSearch_GotFocus()

If PopUpValue1 <> "" Then
   txtSubledger = PopUpValue1
   lblSubledger.Caption = PopUpValue1
   
   txtgenledger.text = PopUpValue2
   txtOpening = PopUpValue3
   txtDescInvoice.text = popupvalue4
   
   txtgenledger.SetFocus
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys "{TAB}"
End Sub

Private Sub Form_Load()
header(0).top = MainMenu.top + 60
header(0).Left = MainMenu.Left
header(0).Width = MainMenu.Width
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False



Set RS = New ADODB.Recordset
RS.Open "select distinct gledger from gledger where " & stringyear & " and slf=1", con, adOpenDynamic, adLockOptimistic, adCmdText
If Not RS.EOF Then
    Do While Not RS.EOF
        txtgenledger.AddItem RS!gledger
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If
RS.Close
    


End Sub

Private Sub Form_Unload(cancel As Integer)
'''MainMenu.Toolbar1.Visible = True
End Sub

Private Sub txtcity_GotFocus()

HIT

If PopUpValue1 <> "" Then
    HIT
    txtCity = PopUpValue2
    txtDistrict = PopUpValue3
    txtState = popupvalue4
    txtbcityid = PopUpValue1
    
    txtPhone.SetFocus
        
End If

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""

End Sub
Sub searchData()
      
  If RS.State = 1 Then RS.Close
  RS.Open "select * from SLEDGER where " & stringyear & " and SUBLEDGER='" & txtSubledger & "'", con
  If RS.EOF = False Then
     
     txtgenledger = RS!gledger
     txtSubledger = RS!subledger
     txtAdd1 = RS!address1
     txtAdd2 = RS!address2
     txtBPhone = RS!phone
     txtFax = RS!FAX
     txtOpening = RS!YEAROPENING
     txtbcityid = RS!bcityid
     txtRAdd1 = RS!radd1
     txtRAdd2 = RS!radd2
     txtRcityid = RS!rcityid
     txtRPhone = RS!rphone
     txtRFax = RS!rfax
     txtEmail = RS!email
     txtTpt = RS!tpt
     txtRange = RS!nRange
     txtBank = RS!bank
     txtRemark = RS!slremark
     txtContact = RS!contperson
     txtBPhone = RS!contphone
     txtTin = RS!tinno
     
     txtCity = PopUpValue3
     txtDistrict = popupvalue4
     txtState = popupvalue5
     
     
     If RS.State = 1 Then RS.Close
     RS.Open "select [state],city from QryMap where " & stringyear & " and cityid='" & txtRcityid & "'"
     If RS.EOF = False Then
       txtRCity = RS(1)
       txtRState = RS(0)
     End If
     
     If RS.State = 1 Then RS.Close
     RS.Open "select [state],city,district from QryMap where " & stringyear & " and cityid='" & txtbcityid & "'"
     If RS.EOF = False Then
       txtCity = RS(1)
       txtState = RS(0)
       txtDistrict = RS(2)
     End If
     
     
     cmdEdit_4.SetFocus
    
     
     
     
'insert into SLEDGER(gledger,subledger,address1,address2,Phone,fax,YEAROPENING,bcityid,radd1,radd2,rcityid,rphone,rfax,
'    email,tpt,nRange,bank,sadd1,slremark,contperson,contphone,tinno,TypeOfCust,FYEAR)
'  values(@gldger,@subledger,@address1,@address2,@Phone,@fax,@YEAROPENING,@bCityId,@radd1,@radd2,@rcityid,@rphone,@rfax,
'    @contmail,@tpt,@range,@bank,@sadd1,@slremark,@contperson,@contphone,@tinno,@TypeOfCust,@FYEAR)
     
     
''CON.Execute "exec insertData_subledger '" & txtgenledger & "','" & txtSubledger & "','" & txtAdd1 & "'," & _
''"'" & txtAdd2 & "','" & txtPhone & "','" & txtFax & "'," & Val(txtOpening) & ",'" & txtbcityid & "'," & _
''"'" & txtRAdd1 & "','" & txtRAdd2 & "','" & (txtRcityid) & "','" & txtPhone & "','" & txtRFax & "'," & _
''"'" & txtEmail & "','" & txtTpt & "','" & txtRange & "','" & txtBank & "','" & txtBAdd1 & "'," & _
''"'" & txtRemark & "','" & txtContact & "','" & txtBPhone & "','" & txtTin & "','" & TypeOfCust & "','" & main.session & "'"

     
     
     
     
  End If
      
End Sub
Private Sub txtcity_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Exit Sub
If KeyCode = 13 Then Exit Sub
tblNo = 6
frmSearchItem.Show

End Sub
Private Sub txtDescInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
 
 If KeyCode = 13 Then
    txtOpening.SetFocus
 End If
 
End Sub

Private Sub txtgenledger_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      sendkeys "{tab}"
   End If
End Sub


Private Sub txtOpening_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then cmdSave_2_Click
End Sub

Private Sub txtRCity_GotFocus()

HIT

If PopUpValue1 <> "" Then
    HIT
    txtRCity = PopUpValue2
    txtRState = PopUpValue3
    txtRcityid = PopUpValue1
    
    txtRPhone.SetFocus
        
End If

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""
End Sub

Private Sub txtRCity_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Exit Sub
If KeyCode = 13 Then Exit Sub
tblNo = 6
frmSearchItem.Show

End Sub
Private Sub txtSubledger_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   sendkeys "{tab}"
End If
End Sub
