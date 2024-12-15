VERSION 5.00
Begin VB.Form screen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3810
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "screen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4140
      Left            =   420
      TabIndex        =   0
      Top             =   0
      Width           =   7260
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   3660
         Top             =   1950
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   2865
         Left            =   420
         Picture         =   "screen.frx":000C
         Stretch         =   -1  'True
         Top             =   750
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1170
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5580
         TabIndex        =   1
         Top             =   810
         Width           =   930
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Devloped By Software Consultant (India)."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   2
         Top             =   570
         Width           =   3900
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Payroll && Invoice Management System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   3
         Top             =   180
         Width           =   6585
      End
   End
End
Attribute VB_Name = "screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Frame1_Click
End Sub

Private Sub Form_Load()
counter = 1
End Sub

Private Sub Frame1_Click()
    Unload Me
    Load MainMenu
    MainMenu.Show
End Sub

Private Sub Timer1_Timer()
If counter > 5 Then
    Timer1.Enabled = False
    Frame1_Click
Else
    counter = counter + 1
End If
End Sub
