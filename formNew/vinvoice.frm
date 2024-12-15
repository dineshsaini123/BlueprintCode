VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form viewinvoice 
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   1815
   ClientWidth     =   14655
   ClipControls    =   0   'False
   Icon            =   "vinvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   14655
   WindowState     =   2  'Maximized
   Begin VB.Frame panel 
      Height          =   9060
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   14865
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   435
         Left            =   8505
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8415
         Width           =   1095
      End
      Begin VB.CommandButton export 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Export"
         Height          =   525
         Left            =   6045
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   8415
         Width           =   705
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "vinvoice.frx":000C
         Left            =   7575
         List            =   "vinvoice.frx":001F
         TabIndex        =   3
         Text            =   "100 %"
         Top             =   8415
         Width           =   855
      End
      Begin VB.CommandButton print 
         BackColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   6765
         Picture         =   "vinvoice.frx":0045
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8415
         Width           =   765
      End
      Begin RichTextLib.RichTextBox r1 
         Height          =   8130
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   14625
         _ExtentX        =   25797
         _ExtentY        =   14340
         _Version        =   393217
         ScrollBars      =   3
         RightMargin     =   20000
         TextRTF         =   $"vinvoice.frx":01B7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComDlg.CommonDialog c1 
         Left            =   5445
         Top             =   8415
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "viewinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLSTRING As String
'Dim CON As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Combo1_Change()
If Trim(Combo1.Text) = "75 %" Then
    r1.Font.Size = 8
    
End If
If Trim(Combo1.Text) = "100 %" Then
    r1.Font.Size = 10
End If
If Trim(Combo1.Text) = "200 %" Then
    r1.Font.Size = 18
  
End If
If Trim(Combo1.Text) = "125 %" Then
    r1.Font.Size = 12
End If
If Trim(Combo1.Text) = "150 %" Then
    r1.Font.Size = 14
End If


End Sub

Private Sub Combo1_Click()
'r1.row = 1
If Trim(Combo1.Text) = "75 %" Then
    r1.Font.Size = 8
End If
If Trim(Combo1.Text) = "100 %" Then
    r1.Font.Size = 10
End If
If Trim(Combo1.Text) = "200 %" Then
    r1.Font.Size = 18
End If
If Trim(Combo1.Text) = "125 %" Then
    r1.Font.Size = 12
End If
If Trim(Combo1.Text) = "150 %" Then
    r1.Font.Size = 14
End If


End Sub

Private Sub Command1_Click()

'   If CRITNOTE.Visible = True Then
'      CRITNOTE.Enabled = True
'   ElseIf INVOICE.Visible = True Then
'      INVOICE.Enabled = True
'   ElseIf countersale.Visible = True Then
'       countersale.Enabled = True
'   End If
'
      
      Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Unload Me
End If
End Sub

Private Sub export_Click()
'    d1.ShowPrinter
'    MsgBox "copies =" + Str(d1.copies)
'    d1.copies
'    Printer.PaperSize
    
End Sub
Public Function printnow()
   Dim X As Long
    Dim p As Printer
    For i = 0 To Printers.Count - 1
        Set p = Printers(i)
        If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.Text) Then
            Exit For
        End If
    Next
    For i = 1 To (Printdlg.UpDown1.value)
       X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt LPT1", vbHide)
    Next
  
End Function
    
Private Sub Form_Load()
 
     
    Set rs = New ADODB.Recordset
    r1.FileName = "" + App.Path + "\vipin.txt"
    r1.LoadFile (r1.FileName)
     
     
'    Me.Width = MainMenu.Width - 2000
'    Me.Height = MainMenu.Height - 4500

BackColorFrom Me
    
End Sub

Private Sub Form_Resize()
    
    'r1.Width = Me.Width - 500
    'r1.Height = Me.Height - 1000
    

    
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30
    
    export.Top = r1.Top + r1.Height + 30
    [print].Top = r1.Top + r1.Height + 30
    
    
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2



End Sub

Private Sub Form_Unload(cancel As Integer)
'   If bankadvice.Visible = True Then
'      bankadvice.Text1.SetFocus
'      bankadvice.Text1.SelStart = 0
'      bankadvice.Text1.SelLength = Len(bankadvice.Text1.Text)
'   End If
End Sub

Private Sub print_Click()
   printnow
End Sub

Private Sub RichTextBox1_Change()
RichTextBox1.w
End Sub
