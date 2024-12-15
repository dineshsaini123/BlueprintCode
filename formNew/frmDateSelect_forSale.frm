VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDateSelect_forSale 
   Caption         =   "Date Selection"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   3600
   Icon            =   "frmDateSelect_forSale.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   690
      Left            =   1395
      TabIndex        =   10
      Top             =   2475
      Width           =   1320
   End
   Begin VB.Frame Frame1_dateselection 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   3030
      Begin MSComCtl2.DTPicker fromDate1 
         Height          =   330
         Left            =   1305
         TabIndex        =   1
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65667073
         CurrentDate     =   39795
      End
      Begin MSComCtl2.DTPicker toDate1 
         Height          =   330
         Left            =   1305
         TabIndex        =   2
         Top             =   540
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65667073
         CurrentDate     =   39795
      End
      Begin MSComCtl2.DTPicker fromDate2 
         Height          =   330
         Left            =   1305
         TabIndex        =   3
         Top             =   1125
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65667073
         CurrentDate     =   39795
      End
      Begin MSComCtl2.DTPicker toDate2 
         Height          =   330
         Left            =   1305
         TabIndex        =   4
         Top             =   1530
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65667073
         CurrentDate     =   39795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   225
         TabIndex        =   9
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   225
         TabIndex        =   8
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Height          =   60
         Left            =   0
         TabIndex        =   7
         Top             =   945
         Width           =   2985
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   225
         TabIndex        =   6
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   225
         TabIndex        =   5
         Top             =   1530
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDateSelect_forSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
'con.Execute "exec netsale 'y'"
End Sub

Private Sub Form_Load()

fromDate1.value = from_date
toDate1.value = to_date

fromDate2.value = from_date
toDate2.value = to_date



End Sub
