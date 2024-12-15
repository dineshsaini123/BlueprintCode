VERSION 5.00
Begin VB.Form SubledgerSearch 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   3960
   ClientTop       =   2415
   ClientWidth     =   10110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6031.782
   ScaleMode       =   0  'User
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4635
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10125
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
         Top             =   480
         Width           =   9945
      End
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
         Height          =   330
         Left            =   0
         TabIndex        =   4
         Top             =   4200
         Visible         =   0   'False
         Width           =   10065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter Text For Search :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1650
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Return"
      Height          =   405
      Left            =   8820
      TabIndex        =   1
      Top             =   5220
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
Text1.Text = Right(FindCombo.Text, 5) & " " & Left(FindCombo.Text, 45)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Dim rs1 As ADODB.Recordset
Me.Top = 0
Me.Left = 0
Set rs1 = New ADODB.Recordset
rs1.Open "Select Subledger,discategory,category2 from sledger where " & stringyear, CON, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then
   rs1.MoveFirst
   While Not rs1.EOF
      If InStr(6, rs1!SUBLEDGER, " ") > 0 Then
         FindCombo.AddItem Trim(Right(Trim(rs1!SUBLEDGER), Len(Trim(rs1!SUBLEDGER)) - 5)) & Space(60 - Len(rs1!SUBLEDGER) - 5) & Left(Trim(rs1!SUBLEDGER), 5) & "    " & rs1!DISCATEGORY & "        " & rs1!category2
      Else
         FindCombo.AddItem rs1!SUBLEDGER
      End If
      rs1.MoveNext
   Wend

 End If
          
BackColorFrom Me
          
End Sub

Private Sub Form_Unload(cancel As Integer)
''MainMenu.Toolbar1.Visible = False
End Sub
