VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAnnexure 
   Caption         =   "Annexure B"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport cr 
      Left            =   375
      Top             =   1950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox txtToDate 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "##/##/####"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   2250
      TabIndex        =   7
      Top             =   1200
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFromDate 
      Height          =   315
      Left            =   2250
      TabIndex        =   6
      Top             =   750
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   3000
      TabIndex        =   5
      Top             =   1950
      Width           =   1440
   End
   Begin VB.CommandButton cmd_Print 
      Caption         =   "Print"
      Height          =   465
      Left            =   1380
      TabIndex        =   4
      Top             =   1950
      Width           =   1365
   End
   Begin VB.ComboBox cboReportPrint 
      Height          =   315
      ItemData        =   "frmAnnexure.frx":0000
      Left            =   1125
      List            =   "frmAnnexure.frx":0016
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   4740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "To : "
      Height          =   195
      Left            =   1500
      TabIndex        =   3
      Top             =   1275
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From : "
      Height          =   195
      Left            =   1500
      TabIndex        =   2
      Top             =   825
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Report By : "
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   225
      Width           =   840
   End
End
Attribute VB_Name = "frmAnnexure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FirmName As String
Public FirmAdd As String
Public FTinNo As String
Public FYear As String
Public ReportType As String
Dim partyname() As String
Dim PAdd() As String
Dim PTin() As String
Dim InvoiceVal() As Integer
Dim InvoiceDate() As String
Dim PName As String
Dim PQty() As Double
Dim TaxVal() As String
Dim TaxCharged() As String
Dim TotalVal() As String
Dim TaxRate() As String
Public Countval As Integer


Private Sub cmd_Exit_Click()
Unload Me
End Sub
Private Sub cmd_Print_Click()



If cboReportPrint.Text = "" Then
ReportType = ""
MsgBox "Please select a report type...", vbCritical
Exit Sub
End If


ReportType = cboReportPrint.Text
If rs.State = 1 Then rs.Close
rs.Open "SELECT [cname], [add1], [add2] FROM Setup", CON, adOpenKeyset, adLockReadOnly
FirmName = rs!cname
FirmAdd = rs!add1
FirmAdd = FirmAdd & rs!add2 & ""
FTinNo = "09178401649"
FYear = "2010-11"
Countval = 0


CON.Execute "DELETE FROM tmpAnnexure"


'============================= Sales Invoice ==============================================================================================

If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM QrySalesAnnexser WHERE TaxType='" & cboReportPrint.Text & "' AND InvoiceDate>=" & _
"convert(smalldatetime,'" & txtFromDate & "',103) AND InvoiceDate <=CONVERT(smalldatetime,'" & _
txtToDate & "',103)", CON, adOpenKeyset, adLockReadOnly
If Not rs.EOF Then

'---------------
ReDim partyname(rs.RecordCount)
ReDim InvoiceVal(rs.RecordCount)
ReDim InvoiceDate(rs.RecordCount)
ReDim PQty(rs.RecordCount)
ReDim TaxVal(rs.RecordCount)
ReDim TaxCharged(rs.RecordCount)
ReDim TotalVal(rs.RecordCount)

For i = 0 To rs.RecordCount - 1

partyname(i) = rs!party
InvoiceVal(i) = rs!INVOICENO
InvoiceDate(i) = rs!InvoiceDate
PQty(i) = Val(rs!TQty)
TaxVal(i) = rs!GAmount
TaxCharged(i) = rs!aexp2am
TotalVal(i) = Val(TaxVal(i)) + Val(TaxCharged(i))

rs.MoveNext
Countval = Countval + 1
Next


'===================================================

ReDim PAdd(Countval)
ReDim PAdd(Countval)
ReDim PTin(Countval)

'===================================================

For i = 0 To Countval - 1
If rs.State = 1 Then rs.Close
If partyname(i) <> "" Then
rs.Open "SELECT [SUBLEDGER],[tinno],[ADDRESS1],[ADDRESS2] FROM [SLEDGER] WHERE Subledger='" & partyname(i) & "'", CON, adOpenKeyset, adLockReadOnly
If Not rs.EOF Then
   
   PAdd(i) = rs!address1
   PAdd(i) = PAdd(i) & rs!ADDRESS2 & ""
   PTin(i) = rs!tinno

End If
End If
Next


ReDim TaxRate(Countval)

For i = 0 To Countval - 1
If rs.State = 1 Then rs.Close
rs.Open "SELECT rate FROM Invoicec WHERE InvoiceNO='" & InvoiceVal(i) & "' AND rate<>'0'", CON, adOpenKeyset, adLockReadOnly
If Not rs.EOF Then
TaxRate(i) = rs!Rate & ""
End If
Next

PName = "Note-book"


If rs.State = 1 Then rs.Close
For i = 0 To Countval - 1
CON.Execute "INSERT INTO tmpAnnexure([FirmName],[FAddress],[TinNo],[FYear],[PartyName],[PAddress],[PTinNO],[InvoiceNo]," & _
"[InvoiceDate],[ProductName],[PQuantity],[TaxableValue],[TaxCharged],[TotalAmount],[TaxRate],[TypeOfReport],[Category]) VALUES('" & _
FirmName & "','" & FirmAdd & "','" & FTinNo & "','" & FYear & "','" & partyname(i) & "','" & PAdd(i) & "','" & _
PTin(i) & "','" & InvoiceVal(i) & "',convert(smalldatetime,'" & InvoiceDate(i) & "',103),'" & PName & "','" & PQty(i) & "','" & _
TaxVal(i) & "','" & TaxCharged(i) & "','" & TotalVal(i) & "','" & TaxRate(i) & "','" & ReportType & "','1')"
Next


'==================================End Code======================================================================================
End If





'============================= Credit Note ==============================================================================================

If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM QryCreditAnnexser WHERE TaxType='" & cboReportPrint.Text & "' AND InvoiceDate>=" & _
"convert(smalldatetime,'" & txtFromDate & "',103) AND InvoiceDate <=CONVERT(smalldatetime,'" & _
txtToDate & "',103)", CON, adOpenKeyset, adLockReadOnly
If Not rs.EOF Then

'---------------
Countval = 0
ReDim partyname(rs.RecordCount)
ReDim InvoiceVal(rs.RecordCount)
ReDim InvoiceDate(rs.RecordCount)
ReDim PQty(rs.RecordCount)
ReDim TaxVal(rs.RecordCount)
ReDim TaxCharged(rs.RecordCount)
ReDim TotalVal(rs.RecordCount)

For i = 0 To rs.RecordCount - 1

partyname(i) = rs!party
InvoiceVal(i) = rs!INVOICENO
InvoiceDate(i) = rs!InvoiceDate
PQty(i) = Val(rs!TQty)
TaxVal(i) = rs!GAmount
TaxCharged(i) = rs!aexp2am
TotalVal(i) = Val(TaxVal(i)) + Val(TaxCharged(i))

rs.MoveNext
Countval = Countval + 1
Next


'===================================================


ReDim PAdd(Countval)
ReDim PAdd(Countval)
ReDim PTin(Countval)

'===================================================

For i = 0 To Countval - 1
If rs.State = 1 Then rs.Close
If partyname(i) <> "" Then
rs.Open "SELECT [SUBLEDGER],[tinno],[ADDRESS1],[ADDRESS2] FROM [SLEDGER] WHERE Subledger='" & partyname(i) & "'", CON, adOpenKeyset, adLockReadOnly
If Not rs.EOF Then
   
   PAdd(i) = rs!address1
   PAdd(i) = PAdd(i) & rs!ADDRESS2 & ""
   PTin(i) = rs!tinno

End If
End If
Next



ReDim TaxRate(Countval)

For i = 0 To Countval - 1
If rs.State = 1 Then rs.Close
rs.Open "SELECT rate FROM Creditc WHERE InvoiceNO='" & InvoiceVal(i) & "' AND rate<>'0'", CON, adOpenKeyset, adLockReadOnly
If Not rs.EOF Then
TaxRate(i) = rs!Rate & ""
End If
Next

PName = "Notebook"


If rs.State = 1 Then rs.Close
For i = 0 To Countval - 1
CON.Execute "INSERT INTO tmpAnnexure([FirmName],[FAddress],[TinNo],[FYear],[PartyName],[PAddress],[PTinNO],[InvoiceNo]," & _
"[InvoiceDate],[ProductName],[PQuantity],[TaxableValue],[TaxCharged],[TotalAmount],[TaxRate],[TypeOfReport],[Category]) VALUES('" & _
FirmName & "','" & FirmAdd & "','" & FTinNo & "','" & FYear & "','" & partyname(i) & "','" & PAdd(i) & "','" & _
PTin(i) & "','" & InvoiceVal(i) & "',convert(smalldatetime,'" & InvoiceDate(i) & "',103),'" & PName & "','" & PQty(i) & "','" & _
TaxVal(i) & "','" & TaxCharged(i) & "','" & TotalVal(i) & "','" & TaxRate(i) & "','" & ReportType & "','2')"
Next


End If

'================================================================================================================================
Dim S_total1, S_total2, S_total3 As Double
Dim C_total1, C_total2, C_total3 As Double

S_total1 = 0
S_total2 = 0
S_total3 = 0

C_total1 = 0
C_total2 = 0
C_total3 = 0



If rs.State = 1 Then rs.Close
rs.Open "select sum(cast(TaxableValue as float)) as Total1 ,sum(cast(TaxCharged as float)) as Total2,sum(cast(TotalAmount as float)) as Total3 FROM [ExportData].[dbo].[tmpAnnexure] where category='1'", CON, adOpenKeyset, adLockReadOnly
If Not IsNull(rs(0)) Then
S_total1 = rs(0)
S_total2 = rs(1)
S_total3 = rs(2)
End If


If rs.State = 1 Then rs.Close
rs.Open "select sum(cast(TaxableValue as float)) as Total1 ,sum(cast(TaxCharged as float)) as Total2,sum(cast(TotalAmount as float)) as Total3 FROM [ExportData].[dbo].[tmpAnnexure] where category='2'", CON, adOpenKeyset, adLockReadOnly
If Not IsNull(rs(0)) Then
C_total1 = rs(0)
C_total2 = rs(1)
C_total3 = rs(2)
End If




CR.Reset
'CR.Connect = "filedsn=chitraExport;Uid=sa;Pwd=sidc"
CR.Connect = constr
CR.ReportFileName = App.Path + "\Reports\SaleTaxReg.rpt"
CR.Formulas(0) = "Total1=" & (S_total1 - C_total1) & ""
CR.Formulas(1) = "Total2=" & (S_total2 - C_total2) & ""
CR.Formulas(2) = "Total3=" & (S_total3 - C_total3) & ""
CR.WindowState = crptMaximized
CR.Action = 1



End Sub

Private Sub Form_Load()
txtFromDate = DateTime.Date
txtToDate = DateTime.Date
'con_open
End Sub
