VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAmountBAmt 
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   Icon            =   "frmAmountBAmt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport cr 
      Left            =   240
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtAmt 
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Top             =   1620
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   2460
      TabIndex        =   5
      Top             =   2340
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   1440
      TabIndex        =   4
      Top             =   2340
      Width           =   975
   End
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   420
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   70778881
      CurrentDate     =   42539
   End
   Begin MSComCtl2.DTPicker ToDate 
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   70778881
      CurrentDate     =   42539
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1620
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   420
      Width           =   1095
   End
End
Attribute VB_Name = "frmAmountBAmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
cr.Reset
cr.ReportFileName = st1 & "\" & directory & "\inv_Cash.rpt"
cr.DataFiles(0) = st1 + "\" + directory & "\data.mdb"
cr.ReplaceSelectionFormula "{Sale_CashList.netamount}>=" & Val(txtAmt) & " and ({Sale_CashList.invoicedate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {Sale_CashList.invoicedate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "'))"
cr.Formulas(0) = "bal1=" & Val(txtAmt) & ""
cr.Formulas(1) = "fdate='" & fromDate.Value & "'"
cr.Formulas(2) = "tdate='" & ToDate.Value & "'"
cr.WindowShowPrintSetupBtn = True
cr.WindowShowPrintBtn = True
cr.WindowState = crptMaximized
cr.Action = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Unload Me
End If
End Sub
Private Sub Form_Load()
    If RS.State = 1 Then RS.Close
    RS.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
    fromDate.Value = RS!yarfrom
    ToDate.Value = RS!yarto
  
  Me.Top = 1200
  Me.Left = 1200
  
End Sub

