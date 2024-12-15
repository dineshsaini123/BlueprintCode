VERSION 5.00
Begin VB.Form SubledgerSearch 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5328.594
   ScaleMode       =   0  'User
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4635
      Left            =   0
      TabIndex        =   3
      Top             =   -90
      Width           =   8385
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   8205
      End
      Begin VB.ComboBox FindCombo 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3780
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   810
         Width           =   8205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter Text For Search :"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   570
         Width           =   1650
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Return"
      Height          =   405
      Left            =   7050
      TabIndex        =   1
      Top             =   4620
      Width           =   1215
   End
End
Attribute VB_Name = "SubledgerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub FindCombo_Click()
Text1.Text = Trim(Mid(FindCombo.Text, 60)) & " " & Trim(Left(FindCombo.Text, 60))

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
rs1.Open "Select Subledger,distcode from sledger where  " & stridnyear & "", CON, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then
   rs1.MoveFirst
   While Not rs1.EOF
      If rs1!distcode <> "" Then
         FindCombo.AddItem Trim(Right(Trim(rs1!subledger), Len(Trim(rs1!subledger)) - InStr(1, Trim(rs1!subledger), " "))) & Space(60 - (Len(rs1!subledger) - InStr(1, Trim(rs1!subledger), " "))) & Left(Trim(rs1!subledger), InStr(1, rs1!subledger, " "))
      Else
         FindCombo.AddItem rs1!subledger
      End If
      rs1.MoveNext
   Wend

 End If
          
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Toolbar1.Visible = True
End Sub
