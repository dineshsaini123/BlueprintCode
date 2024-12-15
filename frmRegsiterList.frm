VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRegsiterList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Register"
   ClientHeight    =   5865
   ClientLeft      =   3600
   ClientTop       =   2535
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8085
   Begin VB.ComboBox cboTax 
      Height          =   315
      ItemData        =   "frmRegsiterList.frx":0000
      Left            =   4320
      List            =   "frmRegsiterList.frx":000D
      TabIndex        =   13
      Top             =   4500
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sale Register (with duty)..."
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4380
      Width           =   2895
   End
   Begin VB.CommandButton cmdSaleRegis 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sale Register (Date Wise)..."
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3420
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Print Sale Register With Amount (Return)"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "P&rint Sale Register With Qty (Return)"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2940
      Width           =   2835
   End
   Begin VB.CommandButton cmdPrin1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "P&rint Sale Register With Qty"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   2895
   End
   Begin VB.ComboBox cboSalesType 
      Height          =   315
      Left            =   2100
      TabIndex        =   6
      Top             =   1560
      Width           =   2445
   End
   Begin Crystal.CrystalReport cr 
      Left            =   6900
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   375
      Left            =   2130
      TabIndex        =   1
      Top             =   780
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Format          =   71761921
      CurrentDate     =   39100
   End
   Begin MSComCtl2.DTPicker fromdate 
      Height          =   375
      Left            =   2130
      TabIndex        =   0
      Top             =   300
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Format          =   71761921
      CurrentDate     =   39100
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   1140
      TabIndex        =   3
      Top             =   5220
      Width           =   5955
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Print Sale Register With Amount"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2370
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      Height          =   1035
      Left            =   960
      Top             =   4080
      Width           =   6195
   End
   Begin VB.Label Label4 
      Caption         =   "Taxable/Non Taxable/All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   4260
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Sale Type :"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1620
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "To Date"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "From Date"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   360
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

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrin1_Click()
If s2 = 1 Then
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_qty.rpt"
    If cboSalesType.Text = "" Then
    cr.SelectionFormula = "{SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    Else
    cr.SelectionFormula = "{SalesRegister.SalesType}='" & cboSalesType.Text & "'  AND {SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    End If
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1

ElseIf s2 = 4 Then
    
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_Exportqty.rpt"
    cr.SelectionFormula = "{ExportSaleRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {ExportSaleRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {ExportSaleRegister.fyear}='" & main.session & "' and {ExportSaleRegister.setupid}=" & main.setupid & " AND {ExportSaleRegister.NetAmount}>0"
'    End If
'    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1


End If

End Sub

Private Sub cmdPrint_Click()
If s2 = 1 Then
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister.rpt"
    If cboSalesType.Text = "" Then
    cr.SelectionFormula = "{SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    Else
    cr.SelectionFormula = "{SalesRegister.SalesType}='" & cboSalesType.Text & "'  AND {SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    End If
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1

ElseIf s2 = 4 Then
    
  cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_Export.rpt"
    cr.SelectionFormula = "{ExportSaleRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {ExportSaleRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {ExportSaleRegister.fyear}='" & main.session & "' and {ExportSaleRegister.setupid}=" & main.setupid & " AND {ExportSaleRegister.NetAmount}>0"
'    End If
'    cr.WindowShowPrintBtn = True
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
'    cr.Reset
'    cr.Connect = constr
'    cr.ReportFileName = strrptpath & "\reports\saleresiter_batchwise.rpt"
'    cr.SelectionFormula = "{Sale_Batchwise.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {Sale_Batchwise.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {Sale_Batchwise.btno}='" & cbobatch.Text & "' and {Sale_Batchwise.fyear}='" & main.session & "' and {Sale_Batchwise.setupid}=" & main.setupid & ""
'    cr.WindowShowPrintBtn = True
'    cr.WindowShowPrintSetupBtn = True
'    cr.Formulas(2) = "fromdate='" & fromdate.Value & "'"
'    cr.Formulas(3) = "todate='" & todate.Value & "'"
'    cr.WindowState = crptMaximized
'    cr.Action = 1
End If
End Sub

Private Sub cmdSaleRegis_Click()
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_DateWise.rpt"
    If cboSalesType.Text = "" Then
    cr.SelectionFormula = "{SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    Else
    cr.SelectionFormula = "{SalesRegister.SalesType}='" & cboSalesType.Text & "'  AND {SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    End If
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1
End Sub

Private Sub Command1_Click()
If s2 = 1 Then
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_qty_ret.rpt"
    If cboSalesType.Text = "" Then
    cr.SelectionFormula = "{SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    Else
    cr.SelectionFormula = "{SalesRegister.SalesType}='" & cboSalesType.Text & "'  AND {SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    End If
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1

ElseIf s2 = 4 Then
    
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_Exportqty.rpt"
    cr.SelectionFormula = "{ExportSaleRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {ExportSaleRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {ExportSaleRegister.fyear}='" & main.session & "' and {ExportSaleRegister.setupid}=" & main.setupid & " AND {ExportSaleRegister.NetAmount}>0"
'    End If
'    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1


End If
End Sub

Private Sub Command2_Click()
If s2 = 1 Then
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_ret.rpt"
    If cboSalesType.Text = "" Then
    cr.SelectionFormula = "{SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    Else
    cr.SelectionFormula = "{SalesRegister.SalesType}='" & cboSalesType.Text & "'  AND {SalesRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {SalesRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {SalesRegister.fyear}='" & main.session & "' and {SalesRegister.setupid}=" & main.setupid & ""
    End If
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1

ElseIf s2 = 4 Then
    
  cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_Export.rpt"
    cr.SelectionFormula = "{ExportSaleRegister.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {ExportSaleRegister.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {ExportSaleRegister.fyear}='" & main.session & "' and {ExportSaleRegister.setupid}=" & main.setupid & " AND {ExportSaleRegister.NetAmount}>0"
'    End If
'    cr.WindowShowPrintBtn = True
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
'    cr.Reset
'    cr.Connect = constr
'    cr.ReportFileName = strrptpath & "\reports\saleresiter_batchwise.rpt"
'    cr.SelectionFormula = "{Sale_Batchwise.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {Sale_Batchwise.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {Sale_Batchwise.btno}='" & cbobatch.Text & "' and {Sale_Batchwise.fyear}='" & main.session & "' and {Sale_Batchwise.setupid}=" & main.setupid & ""
'    cr.WindowShowPrintBtn = True
'    cr.WindowShowPrintSetupBtn = True
'    cr.Formulas(2) = "fromdate='" & fromdate.Value & "'"
'    cr.Formulas(3) = "todate='" & todate.Value & "'"
'    cr.WindowState = crptMaximized
'    cr.Action = 1
End If
End Sub

Private Sub Command3_Click()
    
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_educess.rpt"
    If cboTax.Text = "All" Then
       cr.SelectionFormula = "{salesQry_withEducess.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {salesQry_withEducess.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {salesQry_withEducess.fyear}='" & main.session & "' and {salesQry_withEducess.setupid}=" & main.setupid & ""
    ElseIf cboTax.Text = "Non Taxable" Then
       cr.SelectionFormula = "{salesQry_withEducess.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {salesQry_withEducess.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {salesQry_withEducess.fyear}='" & main.session & "' and {salesQry_withEducess.setupid}=" & main.setupid & " and {salesQry_withEducess.aexp2am}=0"
    ElseIf cboTax.Text = "Taxable" Then
       cr.SelectionFormula = "{salesQry_withEducess.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {salesQry_withEducess.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {salesQry_withEducess.fyear}='" & main.session & "' and {salesQry_withEducess.setupid}=" & main.setupid & " and {salesQry_withEducess.aexp2am}>0"
    End If
    
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.Formulas(2) = "rptheader='" & firm_Address & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1
    

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

cboTax.ListIndex = 0

rs.Open "Select * from setup where " & stringyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
CNSetup
fromdate.Value = rs!yarfrom
todate.Value = rs!yarto
rs.Close

If s2 = 3 Then
Me.Caption = "Purchase Register....."
Else
Me.Caption = "Sale Register....."
End If


If s2 = 1 Then
Command1.Visible = True
Command2.Visible = True
Else
Command1.Visible = False
Command2.Visible = False

End If


If s2 = 1 Or s2 = 3 Then
'cbobatch.Visible = False
'lblbtno.Visible = False

If rs.State = 1 Then rs.Close
    'rs.Open "select subledger from SLEDGER where gledger='SALES' and " & stringyear & "", CON
    rs.Open "SELECT distinct ([SUBLEDGER]) FROM [ExportData].[dbo].[INVOICEC] where " & stringyear, CON
    While rs.EOF = False
    If InStr(rs(0), "%") = 0 Then
        If rs(0) <> "" Then
        cboSalesType.AddItem rs(0)
        End If
    End If
    rs.MoveNext
    Wend

ElseIf s2 = 4 Then
    rs.Open "select subledger from SLEDGER where gledger='SALES RETURN' and " & stringyear & "", CON
    While rs.EOF = False
    cboSalesType.AddItem rs(0)
    rs.MoveNext
    Wend


Else
    If rs.State = 1 Then rs.Close
    rs.Open "select subledger from SLEDGER where gledger='SALES' and " & stringyear & "", CON
    While rs.EOF = False
    cboSalesType.AddItem rs(0)
    rs.MoveNext
    Wend

'cbobatch.Visible = True
'lblbtno.Visible = True
End If

End Sub

Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub todate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

