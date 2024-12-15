VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCash_SalesReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash/Invoice Sales Reports"
   ClientHeight    =   3405
   ClientLeft      =   3375
   ClientTop       =   2625
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5160
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Godown Wise Sales (Qty)"
      Height          =   240
      Left            =   1200
      TabIndex        =   9
      Top             =   540
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Bill Wise Godown Sales"
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Top             =   180
      Width           =   2280
   End
   Begin Crystal.CrystalReport cr 
      Left            =   4320
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72679425
      CurrentDate     =   39506
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   645
      Left            =   2640
      Picture         =   "frmCash_SalesReports.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   645
      Left            =   1140
      Picture         =   "frmCash_SalesReports.frx":0BE4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox cboMarks 
      Height          =   315
      ItemData        =   "frmCash_SalesReports.frx":17C8
      Left            =   2280
      List            =   "frmCash_SalesReports.frx":17D5
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1260
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker toDate 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72679425
      CurrentDate     =   39506
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   255
      Left            =   1140
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      Height          =   255
      Left            =   1140
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Marks"
      Height          =   255
      Left            =   1140
      TabIndex        =   1
      Top             =   1260
      Width           =   1095
   End
End
Attribute VB_Name = "frmCash_SalesReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cash_Click()
'If cash.Value = True Then
'cboMarks.Clear
'If RS.State = 1 Then RS.Close
'RS.Open "select distinct(t2) from casha where  not isnull(t2)", con
'While RS.EOF = False
'cboMarks.AddItem RS(0)
'RS.MoveNext
'Wend
'End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Cmdprint_Click()


CON.Execute "delete from cashandsalesrpt"

CON.Execute "INSERT INTO CashAndSalesrpt (Bill, BillDate,Net,BookName,Qty,Mark,Cash_Sales,gp)  SELECT INVOICEA.INVOICENO, INVOICEA.INVOICEDATE, INVOICEA.NETAMOUNT, BOOKS.BOOKNAME, INVOICEB.QUANTITY, INVOICEA.godown,'INV.',books.GROUPCODE FROM (INVOICEA INNER JOIN INVOICEB ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) INNER JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE where convert(smalldatetime,INVOICEA.INVOICEDATE,103)>=convert(smalldatetime,'" & fromDate.value & "',103) and convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" & toDate.value & "',103) and INVOICEA.godown='" & cboMarks.Text & "'"
CON.Execute "INSERT INTO CashAndSalesrpt (Bill, BillDate,Net,BookName,Qty,Mark,Cash_Sales,gp)  SELECT casha.INVOICENO, casha.INVOICEDATE, casha.NETAMOUNT, BOOKS.BOOKNAME, cashb.QUANTITY, casha.godown,'C/M',books.GROUPCODE FROM (casha INNER JOIN cashb ON casha.INVOICENO = cashb.INVOICENO) INNER JOIN BOOKS ON cashb.BOOKCODE = BOOKS.BOOKCODE where convert(smalldatetime,casha.INVOICEDATE,103)>=convert(smalldatetime,'" & fromDate.value & "',103) and convert(smalldatetime,casha.INVOICEDATE,103)<=convert(smalldatetime,'" & toDate.value & "',103) and cashA.godown='" & cboMarks.Text & "'"


If Me.Option1.value = True Then
If MsgBox("Want to show ?", vbInformation + vbYesNo) = vbYes Then
cr.Reset
cr.ReportFileName = App.Path & "\reports\SalesReports.rpt"
cr.Connect = "filedsn=chitradsn;uid=sa;pwd=sidc;"
'cr.DataFiles(0) = st1 + "\" + Trim(main.directory) & "\data.mdb"
cr.ReplaceSelectionFormula "{templedgerrpt.mark}='" & cboMarks.Text & "' and ({templedgerrpt.billdate}>=datevalue('" & Format(fromDate.value, "MM/dd/yyyy") & "') and {templedgerrpt.billdate}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "'))"
cr.WindowState = crptMaximized
cr.WindowShowRefreshBtn = True
cr.Action = 1
End If
Else
If MsgBox("Want to show ?", vbInformation + vbYesNo) = vbYes Then
   cr.Reset
   cr.ReportFileName = App.Path & "\reports\GodownWiseSalesQty.rpt"
   cr.Connect = "filedsn=chitradsn;uid=sa;pwd=sidc;"
   'cr.DataFiles(0) = st1 + "\" + Trim(main.directory) & "\data.mdb"
   cr.ReplaceSelectionFormula "{templedgerrpt.mark}='" & cboMarks.Text & "' and ({templedgerrpt.billdate}>=datevalue('" & Format(fromDate.value, "MM/dd/yyyy") & "') and {templedgerrpt.billdate}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "'))"
   cr.WindowState = crptMaximized
   cr.WindowShowRefreshBtn = True
   cr.Action = 1
End If

End If

End Sub
Private Sub Form_Load()
If RS.State = 1 Then RS.close
RS.Open "setup1", CON, adOpenDynamic, adLockReadOnly, adCmdTable
If RS.EOF = False Then
fromDate.value = RS!yarfrom
toDate.value = RS!yarto
End If
BackColorFrom Me
End Sub

