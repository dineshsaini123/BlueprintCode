VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form trial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "trial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5790
      TabIndex        =   10
      Top             =   4680
      Width           =   1035
   End
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   1440
      Picture         =   "trial.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4710
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "trial.frx":045D
      Left            =   2370
      List            =   "trial.frx":0470
      TabIndex        =   8
      Text            =   "100 %"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton print 
      Height          =   345
      Left            =   1920
      Picture         =   "trial.frx":0496
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4710
      Width           =   345
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   4095
      Left            =   150
      TabIndex        =   6
      Top             =   150
      Visible         =   0   'False
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7223
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      RightMargin     =   2000
      TextRTF         =   $"trial.frx":0608
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
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   4590
      TabIndex        =   3
      Top             =   3060
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   2400
      TabIndex        =   2
      Top             =   3090
      Width           =   1545
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   3030
      TabIndex        =   0
      Top             =   1560
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      _Version        =   393216
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   5670
      TabIndex        =   1
      Top             =   1500
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      _Version        =   393216
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Caption         =   " - To - "
      Height          =   315
      Left            =   4650
      TabIndex        =   5
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "From The Date"
      Height          =   315
      Left            =   930
      TabIndex        =   4
      Top             =   1590
      Width           =   1995
   End
End
Attribute VB_Name = "trial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Dim rs As Recordset

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        rs.Open "select * from gledger where slf=true where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rs.BOF Then
            rs.Find "gledger='" + Trim(COMBOGENLEDGER.Text) + "'"
            If rs.EOF Then
                COMBOGENLEDGER.SetFocus
            End If
        Else
            COMBOGENLEDGER.SetFocus
        End If
        rs.Close
    End If
End Sub

Private Sub Commandreturn_Click()
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
'    Dim CON As ADODB.Connection
    Dim balance As Double
'    Set CON = New ADODB.Connection
    Set rs1 = New ADODB.Recordset
    balance = 0
'    CON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + Trim(App.Path) + "\" + Trim(main.directory) + "\tchitra.mdb"
'    CON.Open
    If rs.State = 1 Then
        rs.Close
    End If
    rs.Open "SELECT * FROM GLEDGER ORDER BY GLEDGER where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            If rs!SLF = True Then
                rs1.Open "select * from sledger where  " & stringyear & " and gledger='" + Trim(rs!gledger) + "' ", CON, adOpenKeyset, adLockReadOnly, adCmdText
                balance = 0
                If Not rs1.BOF Then
                    Do While Not rs1.EOF
                        balance = balance + myround(rs1!YEAROPENING, 2)
                        rs2.Open "select sum(amount) from vouchers where genledger='" + Trim(rs!gledger) + "' and subledger='" + Trim(rs1!subledger) + "' and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='D'", CON, adOpenKeyset, adLockReadOnly, adCmdText
                        If rs2(0) >= 0 Then
                            balance = balance + rs2(0)
                        End If
                        rs2.Close
                        rs2.Open "select sum(amount) from vouchers where genledger='" + Trim(rs!gledger) + "' and subledger='" + Trim(rs1!subledger) + "' and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='C'", CON, adOpenKeyset, adLockReadOnly, adCmdText
                        If rs2(0) >= 0 Then
                            balance = balance - rs2(0)
                        End If
                        rs2.Close
                        If Not rs1.EOF Then
                            rs1.MoveNext
                        End If
                    Loop
                End If
                rs1.Close
            Else
                balance = 0
                balance = balance + myround(rs1!YEAROPENING, 2)
                rs2.Open "select sum(amount) from vouchers where genledger='" + Trim(rs!gledger) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='D'", CON, adOpenKeyset, adLockReadOnly, adCmdText
                If rs2(0) >= 0 Then
                    balance = balance + rs2(0)
                End If
                rs2.Close
                rs2.Open "select sum(amount) from vouchers where genledger='" + Trim(rs!gledger) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='C'", CON, adOpenKeyset, adLockReadOnly, adCmdText
                If rs2(0) >= 0 Then
                    balance = balance - rs2(0)
                End If
                rs2.Close
            End If
        Loop
    End If
    '****************&&&&&&&&&&&&&&777
Else
    MsgBox "gen. ledger or sub. ledger not selected"
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
Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop
Me.TOP = 0
Me.Left = 0

'Set CON = New ADODB.Connection
Set rs = New ADODB.Recordset
'    With CON
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
'        .Open
'    End With
    rs.Open "select * from gledger where slf=true where " & stringyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            COMBOGENLEDGER.AddItem Trim(rs!gledger)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
    CNSetup
    date1.Text = rs!yarfrom
    date2.Text = rs!yarto
    rs.Close
End Sub

