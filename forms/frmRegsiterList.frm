VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegsiterList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Register"
   ClientHeight    =   2205
   ClientLeft      =   3600
   ClientTop       =   2535
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5055
   Begin VB.ComboBox cbobatch 
      Height          =   315
      Left            =   2130
      TabIndex        =   0
      Top             =   120
      Width           =   1785
   End
   Begin Crystal.CrystalReport cr 
      Left            =   480
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   375
      Left            =   2130
      TabIndex        =   2
      Top             =   1020
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Format          =   55836673
      CurrentDate     =   39100
   End
   Begin MSComCtl2.DTPicker fromdate 
      Height          =   375
      Left            =   2130
      TabIndex        =   1
      Top             =   540
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Format          =   55836673
      CurrentDate     =   39100
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1710
      Width           =   1335
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1710
      Width           =   1455
   End
   Begin VB.Label lblbtno 
      Caption         =   "Batch No"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   150
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "To Date"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "From Date"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmRegsiterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbobatch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
If s2 = 1 Then
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\saleresiter.rpt"
    cr.SelectionFormula = "{saleregister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {saleregister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {saleregister.fyear}='" & main.session & "' and {saleregister.setupid}=" & main.setupid & ""
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1
ElseIf s2 = 3 Then
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\purchaseregister.rpt"
    cr.SelectionFormula = "{saleregister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {saleregister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {saleregister.fyear}='" & main.session & "' and {saleregister.setupid}=" & main.setupid & ""
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1
Else
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\saleresiter_batchwise.rpt"
    cr.SelectionFormula = "{Sale_Batchwise.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {Sale_Batchwise.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {Sale_Batchwise.btno}='" & cbobatch.Text & "' and {Sale_Batchwise.fyear}='" & main.session & "' and {Sale_Batchwise.setupid}=" & main.setupid & ""
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(2) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(3) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1
End If
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

rs.Open "Select * from setup where " & stridnyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
CNSetup
fromdate.Value = rs!yarfrom
todate.Value = rs!yarto
rs.Close

If s2 = 3 Then
Me.Caption = "Purchase Register....."
Else
Me.Caption = "Sale Register....."
End If


If s2 = 1 Or s2 = 3 Then
cbobatch.Visible = False
lblbtno.Visible = False
Else
If rs.State = 1 Then rs.Close
rs.Open "select distinct(btno) from Sale_Batchwise where " & stridnyear & "", CON
While rs.EOF = False
cbobatch.AddItem rs(0)
rs.MoveNext
Wend
cbobatch.Visible = True
lblbtno.Visible = True
End If

End Sub

Private Sub fromdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub todate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

