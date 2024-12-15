VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIssueProduct 
   Caption         =   "Issue Product"
   ClientHeight    =   2688
   ClientLeft      =   4140
   ClientTop       =   2652
   ClientWidth     =   4188
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2688
   ScaleWidth      =   4188
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   1770
      TabIndex        =   0
      Top             =   1530
      Width           =   1575
   End
   Begin Crystal.CrystalReport cr 
      Left            =   930
      Top             =   1530
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker FromDate 
      Height          =   285
      Left            =   1770
      TabIndex        =   1
      Top             =   570
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   487
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   156499971
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   285
      Left            =   1770
      TabIndex        =   2
      Top             =   1050
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   487
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   156499971
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      Height          =   255
      Left            =   450
      TabIndex        =   4
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      Height          =   255
      Left            =   450
      TabIndex        =   3
      Top             =   570
      Width           =   1215
   End
End
Attribute VB_Name = "frmIssueProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
   DSNNew
   cr.Reset
   cr.ReportFileName = rptPath & "\IssueProduct.rpt"
   cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
   cr.ReplaceSelectionFormula "{IssueProducts.dates}>=datevalue('" & Format(FromDate.value, "MM/dd/yy") & "') and {IssueProducts.dates}<=datevalue('" & Format(toDate.value, "MM/dd/yy") & "')"
   cr.Formulas(0) = "fromdate='" & FromDate.value & "'"
   cr.Formulas(1) = "todate='" & toDate.value & "'"
   cr.WindowState = crptMaximized
   cr.Action = 1
End Sub

Private Sub Form_Load()
   FromDate.value = Date
   toDate.value = Date
   Me.BackColor = &HDEFEF8
End Sub



