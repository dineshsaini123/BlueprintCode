VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProductWiseSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Wise Sale"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print Product Wise Sales"
      Height          =   495
      Left            =   900
      TabIndex        =   4
      Top             =   2625
      Width           =   2910
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   900
      TabIndex        =   3
      Top             =   3225
      Width           =   2910
   End
   Begin VB.ComboBox cboSalesType 
      Height          =   315
      Left            =   1275
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   1800
      Width           =   2820
   End
   Begin Crystal.CrystalReport cr 
      Left            =   75
      Top             =   2580
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   375
      Left            =   1965
      TabIndex        =   1
      Top             =   780
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   39100
   End
   Begin MSComCtl2.DTPicker fromdate 
      Height          =   375
      Left            =   1965
      TabIndex        =   2
      Top             =   300
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   39100
   End
   Begin VB.Label Label4 
      Caption         =   "F2  For  Saerch Product"
      Height          =   255
      Left            =   1275
      TabIndex        =   8
      Top             =   1575
      Width           =   2145
   End
   Begin VB.Label Label1 
      Caption         =   "From Date"
      Height          =   255
      Left            =   915
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "To Date"
      Height          =   255
      Left            =   915
      TabIndex        =   6
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Product Wise Sales :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   1260
   End
End
Attribute VB_Name = "frmProductWiseSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbobatch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub
Private Sub cboSalesType_GotFocus()
   HIT
   If PopUpValue1 <> "" Then
      cboSalesType.Text = PopUpValue1
      PopUpValue2 = ""
      PopUpValue1 = ""
      PopUpValue3 = ""
      popupvalue4 = ""
   End If
End Sub

Private Sub cboSalesType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   popuplist10 "select BookNo as Code,TypeofProduct,Rulling,Rate from copymaster where " & stringyear & " order by ProductQuality", CON
   '+' '++ ' ' +rulling+ ' ' +rate+ ' ' +NoofPages
End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub



Private Sub cmdprint_Click()
    
    
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\reports\SaleRegister_Productwise.rpt"
    If cboSalesType.Text = "" Then
    cr.SelectionFormula = "{ProductWiseSale.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {ProductWiseSale.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {ProductWiseSale.fyear}='" & main.session & "' and {ProductWiseSale.setupid}=" & main.setupid & ""
    Else
    cr.SelectionFormula = "{ProductWiseSale.bookcode}='" & cboSalesType.Text & "'  AND {ProductWiseSale.INVOICEDATE}>=datevalue('" & Format(fromdate.Value, "MM/dd/yyyy") & "') and {ProductWiseSale.INVOICEDATE}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {ProductWiseSale.fyear}='" & main.session & "' and {ProductWiseSale.setupid}=" & main.setupid & ""
    End If
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
    cr.Formulas(1) = "todate='" & todate.Value & "'"
    cr.WindowState = crptMaximized
    cr.Action = 1

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

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


If s2 = 1 Or s2 = 3 Then
'cbobatch.Visible = False
'lblbtno.Visible = False

If rs.State = 1 Then rs.Close
    rs.Open "select subledger from SLEDGER where gledger='SALES' and " & stringyear & "", CON
    While rs.EOF = False
    cboSalesType.AddItem rs(0)
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


