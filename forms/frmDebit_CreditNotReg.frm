VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDebit_CreditNotReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debit/Credit Not Register"
   ClientHeight    =   4188
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4848
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4188
   ScaleWidth      =   4848
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Print(Debit Note Register)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   1980
      Width           =   2910
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   3180
      Width           =   2910
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print(Credit Note Register)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2520
      Width           =   2910
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   375
      Left            =   1905
      TabIndex        =   2
      Top             =   840
      Width           =   1785
      _ExtentX        =   3154
      _ExtentY        =   656
      _Version        =   393216
      Format          =   156499969
      CurrentDate     =   39100
   End
   Begin MSComCtl2.DTPicker fromdate 
      Height          =   375
      Left            =   1905
      TabIndex        =   3
      Top             =   360
      Width           =   1785
      _ExtentX        =   3154
      _ExtentY        =   656
      _Version        =   393216
      Format          =   156499969
      CurrentDate     =   39100
   End
   Begin Crystal.CrystalReport CR 
      Left            =   0
      Top             =   300
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label2 
      Caption         =   "To Date"
      Height          =   255
      Left            =   855
      TabIndex        =   5
      Top             =   900
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "From Date"
      Height          =   255
      Left            =   855
      TabIndex        =   4
      Top             =   420
      Width           =   1035
   End
End
Attribute VB_Name = "frmDebit_CreditNotReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
    DSNNew
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = rptPath & "\CreditNotRegister.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.SelectionFormula = "{CreditNotRegister.fyear}='" & main.session & "' and {CreditNotRegister.setupid} in [" & main.setupid & "] and {CreditNotRegister.CND}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {CreditNotRegister.CND}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"
    cr.Formulas(0) = "fromdate='" & FromDate.value & "'"
    cr.Formulas(1) = "todate='" & toDate.value & "'"
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
End Sub

Private Sub Command1_Click()
    DSNNew
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = rptPath & "\DebitNotRegister.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.SelectionFormula = "{DebitNotRegister.fyear}='" & main.session & "' and {DebitNotRegister.setupid} in [" & main.setupid & "] and {DebitNotRegister.DND}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {DebitNotRegister.DND}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"
    cr.Formulas(0) = "fromdate='" & FromDate.value & "'"
    cr.Formulas(1) = "todate='" & toDate.value & "'"
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1

End Sub

Private Sub Form_Load()
 FromDate.value = Date
 toDate.value = Date
End Sub
