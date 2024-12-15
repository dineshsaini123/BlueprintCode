VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog c2 
      Left            =   2940
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   510
      TabIndex        =   0
      Top             =   90
      Width           =   3075
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   1410
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
'&H8&
    c1.ShowPrinter
    If c1.Flags = &H20 Then
     c2.FileName = "c:\kk.txt"
     c2.ShowSave
    Dim kfile  As String
    kfile = c2.FileName
    FileCopy "c:\frunlog.txt", kfile
    Else
        MsgBox "not click"
    End If
End Sub
