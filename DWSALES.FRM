VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form districtwisesales 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "dwsales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   4590
      TabIndex        =   4
      Top             =   4170
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   2430
      TabIndex        =   3
      Top             =   4170
      Width           =   1545
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   3030
      TabIndex        =   1
      Top             =   2400
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      _Version        =   393216
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   3090
      TabIndex        =   0
      Top             =   1140
      Width           =   3885
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   5490
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      _Version        =   393216
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   315
      Left            =   4590
      TabIndex        =   7
      Top             =   2460
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "From The Date"
      Height          =   315
      Left            =   930
      TabIndex        =   6
      Top             =   2400
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Account Desc."
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   1170
      Width           =   2055
   End
End
Attribute VB_Name = "districtwisesales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As Connection
Dim rs As Recordset

Private Sub COMBOGENLEDGER_Change()
COMBOGENLEDGER = UCase(COMBOGENLEDGER)
    If rs.State = 1 Then
        rs.Close
    End If
    rs.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'    Combosubledger.Clear
    If Not rs.BOF Then
        Do While Not rs.EOF
            Combosubledger.AddItem Trim(rs!subledger)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    
End Sub
Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        rs.Open "select * from gledger where slf=FALSE", con, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.BOF Then
                COMBOGENLEDGER.SetFocus
        End If
        rs.Close
    End If
End Sub

Private Sub Commandreturn_Click()
MainMenu.Toolbar1.Visible = True
    Unload Me
End Sub

Private Sub Commandshow_Click()
If Trim(COMBOGENLEDGER.Text) <> "" Then
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
    End If
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim CON1 As ADODB.Connection
    Set CON1 = New ADODB.Connection
    Set rs2 = New ADODB.Recordset
    Dim balance As Double
    balance = 0
    CON1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + Trim(VB.App.Path) + "\" + Trim(main.directory) + "\tchitra.mdb"
    CON1.Open
    If rs.State = 1 Then
        rs.Close
    End If
    rs.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "' order by distcode", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            
            
            
            
            
            
            
            
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    viewledger.Show
Else
    MsgBox "gen. ledger not selected"
End If
End Sub

Private Sub date1_LostFocus()
    If Trim(date1.Text) <> "" Then
        If Not checkdate(Trim(date1.Text), date1) Then
            date1.SetFocus
        End If
    End If
End Sub
Private Sub date2_LostFocus()
    If Trim(date2.Text) <> "" Then
        If Not checkdate(Trim(date2.Text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
Do While Trim(UCase(VB.screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
    If Trim(UCase(VB.screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.screen.ActiveForm
    End If
Loop
Me.Top = 0
Me.Left = 0
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
    With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
        .Open
    End With
    rs.Open "select * from gledger where slf=FALSE", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            COMBOGENLEDGER.AddItem Trim(rs!gledger)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "setup1", con, adOpenKeyset, adLockReadOnly, adCmdTable
    date1.Text = rs!yarfrom
    date2.Text = rs!yarto
    rs.Close
End Sub
