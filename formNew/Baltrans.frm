VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Opening Transfer Option"
   ClientHeight    =   3312
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6936
   Icon            =   "Baltrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3312
   ScaleWidth      =   6936
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox txtsource 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Baltrans.frx":000C
      Left            =   1320
      List            =   "Baltrans.frx":000E
      TabIndex        =   0
      Top             =   1140
      Width           =   4425
   End
   Begin MSComDlg.CommonDialog cm 
      Left            =   6000
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select Path"
      Height          =   420
      Left            =   5880
      TabIndex        =   7
      Top             =   1140
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      ItemData        =   "Baltrans.frx":0010
      Left            =   1320
      List            =   "Baltrans.frx":0012
      TabIndex        =   1
      Top             =   1620
      Width           =   4425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   630
      Left            =   1320
      TabIndex        =   2
      Top             =   2025
      Width           =   4365
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Path"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gen. Ledger"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1740
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pls Run from ledger - goto opening tab and press the closing "
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   420
      Width           =   6885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pls Run Subledger trial for the old financial year for that particular ledger"
      Height          =   270
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   7845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If MsgBox("Want to Transfer", vbCritical + vbYesNo) = vbNo Then
Exit Sub
End If

Screen.MousePointer = vbHourglass

Dim con_dest As New ADODB.Connection
Dim ss As String
Dim op As Double
Dim rs_source As ADODB.Recordset
Dim rs_dest As New ADODB.Recordset
Dim sum1 As Double


Dim db, db1 As String
sum1 = 0
db = ""
db1 = ""


If txtsource = "2016-17" Then
   db1 = "1617"
ElseIf txtsource = "2017-18" Then
   db1 = "1718"
ElseIf txtsource = "2018-19" Then
   db1 = "1819"
ElseIf txtsource = "2019-20" Then
   db1 = "1920"
ElseIf txtsource = "2020-21" Then
   db1 = "2021"
ElseIf txtsource = "2021-22" Then
   db1 = "2122"
ElseIf txtsource = "2022-23" Then
   db1 = "2223"
ElseIf txtsource = "2023-24" Then
   db1 = "2324"
ElseIf txtsource = "2024-25" Then
   db1 = "2425"
   
End If

db = "chitradata_" & db1

Set con_dest = New ADODB.Connection
If LCase(server_) = "server" Then
   con_dest.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db & "; UID=" & sql_user & "; PWD=" & sql_pass
   con_dest.Open
Else
   con_dest.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=COMPAQ\SQL2008NEW; DATABASE=" & db & "; UID=; PWD=;"
   con_dest.Open
End If

con_dest.Execute "update sledger set YEAROPENING=0  where gledger = '" & Me.COMBOGENLEDGER.text & "' "

Set rs_source = New ADODB.Recordset
If rs_source.State = 1 Then rs_source.Close
rs_source.Open "select SUBLEDGER from SLEDGER where gledger = '" & Me.COMBOGENLEDGER.text & "' group by SUBLEDGER", con
Do While Not rs_source.EOF

If RS.State = 1 Then RS.Close
RS.Open "select  sum(OpeningBalance),sum(DAmount),sum(CAmount) from TemprptTrialBalance1 where LEFT(SUBLEDGER,5) = '" + Left(rs_source!subledger, 5) + "'", con
   sum1 = IIf(IsNull(RS(0)), 0, RS(0)) + IIf(IsNull(RS(1)), 0, RS(1)) - Abs(IIf(IsNull(RS(2)), 0, RS(2)))
   ssss = Left(rs_source!subledger, 5)
       
       'If ssss = "M1012" Then
       'MsgBox "ss"
       'End If
       
       
       DoEvents
       DoEvents
       DoEvents
       
       con_dest.Execute "update sledger set YEAROPENING=" & sum1 & " where gledger = '" & Me.COMBOGENLEDGER.text & "' AND LEFT(SUBLEDGER,5) = '" + Left(rs_source!subledger, 5) + "'"
           If Not rs_source.EOF Then
              rs_source.MoveNext
       End If
       DoEvents
       DoEvents
Loop

MsgBox "Done ...", vbInformation
con_dest.Close

Screen.MousePointer = vbDefault


End Sub
Private Sub Command2_Click()
    cm.ShowOpen
    Me.txtsource.text = cm.filename
End Sub
Private Sub Form_Load()
    d1 = 0
    'COMBOGENLEDGER.AddItem "SUNDRY DEBTORS"
    'COMBOGENLEDGER.AddItem "SUNDRY CREDITORS"
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select gledger from GLEDGER where SLF=1", con
    While rs1.EOF = False
    COMBOGENLEDGER.AddItem rs1(0)
    rs1.MoveNext
    Wend
        
        
    txtsource.Clear
    
'    If RS.State = 1 Then RS.close
'    RS.Open "select fyear from financialyear", CCON
'    While RS.EOF = False
'      txtsource.AddItem RS(0)
'      d1 = d1 + 1
'      RS.MoveNext
'    Wend
    
    If main.session = "2018-19" Then
       txtsource.AddItem "2019-20"
    ElseIf main.session = "2019-20" Then
       txtsource.AddItem "2020-21"
    ElseIf main.session = "2020-21" Then
       txtsource.AddItem "2021-22"
    ElseIf main.session = "2021-22" Then
       txtsource.AddItem "2022-23"
    ElseIf main.session = "2022-23" Then
       txtsource.AddItem "2023-24"
    ElseIf main.session = "2023-24" Then
       txtsource.AddItem "2024-25"
       
    ElseIf main.session = "2017-18" Then
       txtsource.AddItem "2018-19"
    ElseIf main.session = "2016-17" Then
       txtsource.AddItem "2017-18"
    ElseIf main.session = "2015-16" Then
       txtsource.AddItem "2016-17"
    End If
    
    txtsource.ListIndex = 0
    
    
End Sub
