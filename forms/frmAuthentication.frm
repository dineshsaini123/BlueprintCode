VERSION 5.00
Begin VB.Form frmAuthentication 
   Caption         =   "Authentication"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   8025
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   10995
   End
   Begin VB.Label Label1 
      Caption         =   "Bill List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4395
   End
End
Attribute VB_Name = "frmAuthentication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
