VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsaleRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Register"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   330
      Left            =   1245
      TabIndex        =   4
      Top             =   1725
      Width           =   1680
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   1275
      TabIndex        =   1
      Top             =   1005
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20578305
      CurrentDate     =   39099
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1260
      TabIndex        =   0
      Top             =   585
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   529
      _Version        =   393216
      Format          =   20578305
      CurrentDate     =   39099
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   285
      Left            =   210
      TabIndex        =   3
      Top             =   1020
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "From "
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   570
      Width           =   870
   End
End
Attribute VB_Name = "frmsaleRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
