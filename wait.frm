VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Waitwindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wait Please.........."
   ClientHeight    =   900
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   531.749
   ScaleMode       =   0  'User
   ScaleWidth      =   5309.739
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox label1 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2250
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Waitwindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean
Private Sub cmdCancel_Click()
'set the global var to false
'to denote a failed login
LoginSucceeded = False
Me.Hide
End Sub
Private Sub cmdOk_Click()
'check for correct password
'For k1 = 1 To 23
'Me.pb1.ma
'Next
End Sub

