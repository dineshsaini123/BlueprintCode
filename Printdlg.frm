VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Printdlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printer Setting"
   ClientHeight    =   4140
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog c1 
      Left            =   1500
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   3660
      Width           =   1125
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4620
      TabIndex        =   16
      Top             =   3660
      Width           =   1245
   End
   Begin VB.Frame frame 
      Caption         =   "Printer"
      Height          =   1515
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7215
      Begin VB.ComboBox Comboprinterstype 
         Height          =   315
         Left            =   1380
         TabIndex        =   18
         Top             =   720
         Width           =   2565
      End
      Begin VB.ComboBox Comboprinters 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   270
         Width           =   2565
      End
      Begin VB.Label ptype 
         Caption         =   "Type"
         Height          =   225
         Left            =   1380
         TabIndex        =   15
         Top             =   750
         Width           =   3075
      End
      Begin VB.Label Label5 
         Caption         =   "Where :"
         Height          =   315
         Left            =   420
         TabIndex        =   14
         Top             =   750
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "&Name :"
         Height          =   315
         Left            =   420
         TabIndex        =   2
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Copies"
      Height          =   1755
      Left            =   3660
      TabIndex        =   10
      Top             =   1590
      Width           =   3585
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   2130
         TabIndex        =   12
         Top             =   390
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text1"
         BuddyDispid     =   196617
         OrigLeft        =   1770
         OrigTop         =   1470
         OrigRight       =   2010
         OrigBottom      =   1935
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1470
         TabIndex        =   11
         Text            =   "1"
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "No. of Copies"
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   420
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Page range"
      Height          =   1755
      Left            =   30
      TabIndex        =   3
      Top             =   1590
      Width           =   3585
      Begin VB.TextBox to 
         Height          =   315
         Left            =   2730
         TabIndex        =   9
         Top             =   630
         Width           =   555
      End
      Begin VB.TextBox from 
         Height          =   315
         Left            =   1620
         TabIndex        =   7
         Top             =   630
         Width           =   555
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pa&ges :"
         Height          =   285
         Left            =   180
         TabIndex        =   5
         Top             =   630
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&All :"
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   690
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "From"
         Height          =   255
         Left            =   1170
         TabIndex        =   6
         Top             =   690
         Width           =   405
      End
   End
End
Attribute VB_Name = "Printdlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Option Explicit

Private Sub CancelButton_Click()
    
    'c1.PrinterDefault = False
    c1.FileName = "c:\chitra\vipin.txt"
    'c1.CancelError = False
    Unload Me
    c1.Flags = &H8&
    c1.ShowPrinter
End Sub
Private Sub Comboprinters_Change()
    Dim p As Printer
    For I = 0 To Printers.Count - 1
        Set p = Printers(I)
        If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.Text) Then
            Exit For
        End If
    Next
    ptype.Caption = p.Port
End Sub
Private Sub Comboprinters_Click()
Dim p As Printer
    For I = 0 To Printers.Count - 1
        Set p = Printers(I)
        If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.Text) Then
            Exit For
        End If
    Next
    ptype.Caption = p.Port
End Sub

Private Sub Form_Load()
    Dim p As Printer
    Dim I As Integer
    For I = 0 To Printers.Count - 1
        Set p = Printers(I)
        Comboprinters.AddItem p.DeviceName
    Next
    Comboprinterstype.AddItem "Lpt1"
    Comboprinterstype.AddItem "File"
    'Comboprinters.Text = Printer.DeviceName
    ptype.Caption = Printer.Port
    Me.Top = (VB.Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (VB.Screen.Width / 2) - (Me.Width / 2)
End Sub
Private Sub Label6_Click()
End Sub
Private Sub OKButton_Click()
    printnow
    Me.Hide
End Sub
Public Function printnow()
    Dim X As Long
    Dim p As Printer
    For I = 0 To Printers.Count - 1
        Set p = Printers(I)
        If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.Text) Then
            Exit For
        End If
    Next
    For I = 1 To (Printdlg.UpDown1.value)
        X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(p.Port), vbHide)
    Next
    Printdlg.UpDown1.value = 1
    'Printdlg.Text1.TEXT = "1"1
End Function

