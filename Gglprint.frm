VERSION 5.00
Begin VB.Form Genledgerprinting 
   Caption         =   "General ledger Printing"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   7350
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Select Order of Printing"
      Height          =   2865
      Left            =   0
      TabIndex        =   12
      Top             =   3120
      Width           =   9465
      Begin VB.OptionButton Ascending 
         Caption         =   "&Ascending"
         Height          =   525
         Left            =   120
         TabIndex        =   17
         Top             =   1650
         Width           =   1305
      End
      Begin VB.OptionButton Descending 
         Caption         =   "&Descending"
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   2190
         Width           =   1305
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Gglprint.frx":0000
         Left            =   150
         List            =   "Gglprint.frx":000D
         TabIndex        =   15
         Text            =   "GENERAL LEDGER"
         Top             =   1290
         Width           =   2445
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Show Before Printing"
         Height          =   435
         Left            =   4680
         TabIndex        =   14
         Top             =   690
         Width           =   3165
      End
      Begin VB.CommandButton Print 
         Caption         =   "Print"
         Height          =   555
         Left            =   4620
         TabIndex        =   13
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Sorted On."
         Height          =   345
         Left            =   150
         TabIndex        =   18
         Top             =   810
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      Begin VB.ComboBox result3 
         Height          =   315
         ItemData        =   "Gglprint.frx":0039
         Left            =   5520
         List            =   "Gglprint.frx":0043
         TabIndex        =   11
         Top             =   1290
         Width           =   2955
      End
      Begin VB.ComboBox result2 
         Height          =   315
         ItemData        =   "Gglprint.frx":0054
         Left            =   5520
         List            =   "Gglprint.frx":005E
         TabIndex        =   10
         Top             =   870
         Width           =   2955
      End
      Begin VB.ComboBox check3 
         Height          =   315
         ItemData        =   "Gglprint.frx":006F
         Left            =   1740
         List            =   "Gglprint.frx":0085
         TabIndex        =   9
         Top             =   1290
         Width           =   2445
      End
      Begin VB.ComboBox condition3 
         Height          =   315
         ItemData        =   "Gglprint.frx":00CB
         Left            =   4260
         List            =   "Gglprint.frx":00E4
         TabIndex        =   8
         Text            =   "="
         Top             =   1290
         Width           =   1185
      End
      Begin VB.ComboBox andor2 
         Height          =   315
         ItemData        =   "Gglprint.frx":0103
         Left            =   630
         List            =   "Gglprint.frx":0110
         TabIndex        =   7
         Top             =   1290
         Width           =   1065
      End
      Begin VB.ComboBox check2 
         Height          =   315
         ItemData        =   "Gglprint.frx":011F
         Left            =   1740
         List            =   "Gglprint.frx":0135
         TabIndex        =   6
         Top             =   870
         Width           =   2475
      End
      Begin VB.ComboBox condition2 
         Height          =   315
         ItemData        =   "Gglprint.frx":017B
         Left            =   4260
         List            =   "Gglprint.frx":0194
         TabIndex        =   5
         Text            =   "="
         Top             =   870
         Width           =   1185
      End
      Begin VB.ComboBox andor1 
         Height          =   315
         ItemData        =   "Gglprint.frx":01B3
         Left            =   630
         List            =   "Gglprint.frx":01C0
         TabIndex        =   4
         Top             =   870
         Width           =   1065
      End
      Begin VB.ComboBox result1 
         Height          =   315
         ItemData        =   "Gglprint.frx":01CF
         Left            =   5520
         List            =   "Gglprint.frx":01D9
         TabIndex        =   3
         Top             =   450
         Width           =   2955
      End
      Begin VB.ComboBox condition1 
         Height          =   315
         ItemData        =   "Gglprint.frx":01EA
         Left            =   4260
         List            =   "Gglprint.frx":0203
         TabIndex        =   2
         Text            =   "="
         Top             =   450
         Width           =   1185
      End
      Begin VB.ComboBox check1 
         Height          =   315
         ItemData        =   "Gglprint.frx":0222
         Left            =   1770
         List            =   "Gglprint.frx":0238
         TabIndex        =   1
         Top             =   450
         Width           =   2445
      End
   End
End
Attribute VB_Name = "Genledgerprinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim con2 As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim stri As String
Dim I, Y As Integer
Dim T1, T2, T3 As String

Private Sub print_Click()
    Dim SQL As String
''    Set CON = New ADODB.Connection
'    With CON
'    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
'    .Open
'    End With
    Set rs1 = New ADODB.Recordset
    
    SQL = "select * from gledger where " & stridnyear & " "
If check1.Text <> "" And Me.condition1.Text <> "" And result1.Text <> "" Then
    SQL = SQL + " where "
    If Trim(check1.Text) = "GENERAL LEDGER" Then
       SQL = SQL + "gledger"
    End If
    If Trim(check1.Text) = "CATEGORY" Then
       SQL = SQL + "CATEGORY"
    End If
    If Trim(check1.Text) = "PLC" Then
       SQL = SQL + "PLC"
    End If
    If Trim(check1.Text) = "BSC" Then
       SQL = SQL + "BSC"
    End If
    If Trim(check1.Text) = "CON. SUBLEDGER" Then
       SQL = SQL + "slf"
    End If
    If Trim(check1.Text) = "YEAR OPENING" Then
       SQL = SQL + "YEAROPENING"
    End If
    
    SQL = SQL + " " + Trim(condition1.Text) + " "
    If Trim(check1.Text) = "GENERAL LEDGER" Or Trim(check1.Text) = "CATEGORY" Then
        SQL = SQL + "'" + Trim(result1.Text) + "'"
    Else
        If Trim(check1.Text) = "YEAR OPENING" Then
            SQL = SQL + Trim(Str(Val(Trim(result1.Text))))
        Else
            SQL = SQL + Trim(result1.Text)
        End If
    End If
End If
MsgBox SQL
a = printreport(SQL, "aa")
'rs1.Open SQL, CON, adOpenKeyset, adLockReadOnly, adCmdText
'If Not rs1.BOF Then
'    rs1.MoveLast
'    Y = rs1.RecordCount
'    rs1.MoveFirst
'End If
''////////////////*********************
'Dim line As Integer
'line = 0
'If Not rs1.BOF Then
'        Open "c:\chitra\vipin.txt" For Output As #1
'        If Not rs1.EOF Then
'            Print #1, Space(2); Chr(27) + Chr(15); lsets("CATEGORY", rs1(0).DefinedSize); lsets("GEN. LEDGER NAME", rs1(1).DefinedSize - 5); lsets("PLC", 6); lsets("BSC", 6); lsets("SLF", 6); lsets("YEAROPENING", 13)
'            Print #1, "----------------------------------------------------------------------------"
'            line = 2
'        End If
'
'            Waitwindow.pb1.Min = 0
'            Waitwindow.pb1.Value = 0
'            Waitwindow.pb1.Max = Y
'            Do While Not rs1.EOF
'
'                Print #1, Space(2); lsets(rs1(0), rs1(0).DefinedSize); lsets(rs1(1), rs1(1).DefinedSize); bsets(rs1(2)); bsets(rs1(3)); bsets(rs1(4)); rsets(rs1(5), rs1(5).DefinedSize)
'                line = line + 1
'                Waitwindow.pb1.Value = Waitwindow.pb1.Value + 1
'                If Not rs1.EOF Then
'                    rs1.MoveNext
'                End If
'            Loop
'            Waitwindow.Hide
'            If line < 80 Then
'                Do While Not line = 80
'                    Print #1, " "
'                    line = line + 1
'                Loop
'            End If
'            Close #1
'            GRIDpreview.SQL (SQL)
'End If
    
End Sub
