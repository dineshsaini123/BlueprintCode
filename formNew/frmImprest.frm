VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImprest 
   Caption         =   "Imprest Register..."
   ClientHeight    =   2688
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   6480
   Icon            =   "frmImprest.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2688
   ScaleWidth      =   6480
   Begin Crystal.CrystalReport cr 
      Left            =   180
      Top             =   2115
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cboagent 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmImprest.frx":000C
      Left            =   1395
      List            =   "frmImprest.frx":000E
      TabIndex        =   2
      Top             =   720
      Width           =   4935
   End
   Begin VB.CommandButton CommandReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   645
      Left            =   2610
      Picture         =   "frmImprest.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1485
      Width           =   1185
   End
   Begin VB.CommandButton CommandPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   645
      Left            =   1395
      Picture         =   "frmImprest.frx":0BF4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1485
      Width           =   1185
   End
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   315
      Left            =   1425
      TabIndex        =   3
      Top             =   225
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      Format          =   154075137
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txtto 
      Height          =   315
      Left            =   3285
      TabIndex        =   4
      Top             =   225
      Width           =   1365
      _ExtentX        =   2413
      _ExtentY        =   550
      _Version        =   393216
      Format          =   154075137
      CurrentDate     =   42409
   End
   Begin VB.Label lblAson 
      Caption         =   "To"
      Height          =   255
      Left            =   2925
      TabIndex        =   6
      Top             =   225
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Imprest Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   90
      TabIndex        =   5
      Top             =   750
      Width           =   1425
   End
End
Attribute VB_Name = "frmImprest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandPrint_Click()

cr.Reset
cr.ReportFileName = rptPath & "/ImprestRegister.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass

If cboagent.Text <> "" Then
   cr.ReplaceSelectionFormula "{tmpbook.subledger}='" & cboagent.Text & "' and ({tmpbook.billdate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {tmpbook.billdate}<=datevalue('" & Format(txtto.value, "MM/dd/yyyy") & "'))"
Else
   cr.ReplaceSelectionFormula "({tmpbook.billdate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {tmpbook.billdate}<=datevalue('" & Format(txtto.value, "MM/dd/yyyy") & "'))"
End If

cr.Formulas(0) = "fdate='" & txtFrom.value & "'"
cr.Formulas(1) = "tdate='" & txtto.value & "'"

cr.WindowShowPrintSetupBtn = True
cr.WindowShowRefreshBtn = True
cr.WindowMaxButton = True
cr.WindowState = crptMaximized
cr.Action = 1

End Sub

Private Sub CommandReturn_Click()
Unload Me
End Sub
Private Sub Form_Load()

Me.Top = 100
Me.Left = 100
Me.Width = 6500
Me.Height = 3345

txtFrom.value = from_date
txtto.value = to_date

If RS.State = 1 Then RS.close
RS.Open "select SubLedger from ImprestBillRegister group by SubLedger", con
While RS.EOF = False
cboagent.AddItem RS(0)
RS.MoveNext
Wend


BackColorFrom Me
End Sub
