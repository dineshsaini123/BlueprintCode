VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWeatorWiseSale 
   Caption         =   "REPORTS"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboItemGp 
      Height          =   315
      Left            =   4680
      TabIndex        =   30
      Top             =   900
      Width           =   3135
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Repair Reports"
      Height          =   240
      Left            =   12000
      TabIndex        =   25
      Top             =   5475
      Visible         =   0   'False
      Width           =   375
      Begin VB.CommandButton Command13 
         Caption         =   "Sites Wise Pending Repair "
         Height          =   495
         Left            =   90
         TabIndex        =   29
         Top             =   1980
         Width           =   2715
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Item Wise Supply"
         Height          =   495
         Left            =   90
         TabIndex        =   28
         Top             =   1440
         Width           =   2715
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Supplier Wise Item Wise Supply"
         Height          =   495
         Left            =   90
         TabIndex        =   27
         Top             =   900
         Width           =   2715
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Item Wise Pending Repair "
         Height          =   495
         Left            =   90
         TabIndex        =   26
         Top             =   360
         Width           =   2715
      End
   End
   Begin VB.ComboBox cboSupp 
      Height          =   315
      Left            =   4680
      TabIndex        =   23
      Top             =   1620
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchse Reports"
      Height          =   3315
      Left            =   6900
      TabIndex        =   19
      Top             =   2295
      Width           =   2925
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Purchase Report"
         Height          =   495
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   2715
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Supplier Wise Item Wise Supply"
         Height          =   495
         Left            =   90
         TabIndex        =   21
         Top             =   900
         Width           =   2715
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Item Wise Supply"
         Height          =   495
         Left            =   90
         TabIndex        =   20
         Top             =   1440
         Width           =   2715
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&EXIT"
      Height          =   495
      Left            =   1005
      TabIndex        =   18
      Top             =   5670
      Width           =   8835
   End
   Begin VB.ComboBox cboItem 
      Height          =   315
      Left            =   4680
      TabIndex        =   14
      Top             =   1260
      Width           =   3135
   End
   Begin VB.ComboBox cboDeptt 
      Height          =   315
      Left            =   4680
      TabIndex        =   13
      Top             =   540
      Width           =   3135
   End
   Begin VB.ComboBox cboCollege 
      Height          =   315
      Left            =   4680
      TabIndex        =   12
      Top             =   180
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   960
      TabIndex        =   7
      Top             =   2385
      Width           =   3000
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Raw Stock Summary"
         Height          =   495
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   2805
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Daily Cunsumption Summary"
         Height          =   495
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1485
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Item Wise for All Deptt"
         Height          =   495
         Left            =   135
         TabIndex        =   8
         Top             =   2100
         Visible         =   0   'False
         Width           =   2805
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Issue Reports"
      Height          =   3315
      Left            =   3930
      TabIndex        =   4
      Top             =   2340
      Width           =   3000
      Begin VB.CommandButton Command5 
         Caption         =   "&Item Wise for All Deppartment"
         Height          =   495
         Left            =   90
         TabIndex        =   11
         Top             =   1440
         Width           =   2805
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Deppartment  Wise  Issue Report"
         Height          =   495
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   900
         Width           =   2805
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Group/Item Wise Issue Report"
         Height          =   495
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   2805
      End
   End
   Begin Crystal.CrystalReport cr 
      Left            =   240
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker FromDate 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   180
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   65994755
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   540
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   65994755
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Group"
      Height          =   255
      Left            =   3555
      TabIndex        =   31
      Top             =   960
      Width           =   1110
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      Height          =   255
      Left            =   3540
      TabIndex        =   24
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      Height          =   255
      Left            =   3540
      TabIndex        =   17
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Deppartment"
      Height          =   255
      Left            =   3540
      TabIndex        =   16
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sites"
      Height          =   255
      Left            =   3540
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      Height          =   255
      Left            =   780
      TabIndex        =   3
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      Height          =   255
      Left            =   780
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmWeatorWiseSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Private Sub cbostaff_Click()
  On Error Resume Next
  Dim ss As New ADODB.Recordset
  If ss.State = 1 Then ss.Close
  ss.Open "select BrokerName from staff where Brokercode=" & cbostaff.Text & "", CON
  If ss.EOF = False Then
     name1.Caption = ss.Fields(0).Value
  End If

End Sub
Private Sub cbodeptt_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      FromDate.SetFocus
   End If
End Sub

Private Sub cmdcr_Click()
   CR.Reset
   CR.ReportFileName = App.Path & "\purchasereg_cr.rpt"
   CR.Connect = constr
   If cboDeptt.Text <> "" Then
    CR.ReplaceSelectionFormula "{finishpurchase.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {finishpurchase.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yy") & "') and {finishpurchase.itemname}='" & cboDeptt.Text & "' and {finishpurchase.credit}=true"
   Else
    CR.ReplaceSelectionFormula "{finishpurchase.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {finishpurchase.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yy") & "') and {finishpurchase.credit}=true"
   End If
   
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.Action = 1
End Sub
Sub SetCondition()
   s = ""
   s = "{Demand.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {Demand.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yy") & "')"
    If cboCollege.Text <> "" Then
      s = s & " and " & "{demand.Supplier}='" & cboCollege.Text & "'"
   End If
   If cboDeptt.Text <> "" Then
      s = s & " and " & "{demand.Deppt}='" & cboDeptt.Text & "'"
   End If
   If cboItem.Text <> "" Then
      s = s & " and " & "{demand.ItemName}='" & cboItem.Text & "'"
   End If
End Sub
Sub SetIssueItemCondition()
   s = ""
   s = "{IssueDeppt.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {IssueDeppt.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yy") & "')"
    If cboCollege.Text <> "" Then
      s = s & " and " & "{IssueDeppt.Supplier}='" & cboCollege.Text & "'"
   End If
   If cboDeptt.Text <> "" Then
      s = s & " and " & "{IssueDeppt.Deppt}='" & cboDeptt.Text & "'"
   End If
   If cboItem.Text <> "" Then
      s = s & " and " & "{IssueDeppt.ItemName}='" & cboItem.Text & "'"
   End If
   
   If cboItemGp.Text <> "" Then
      s = s & " and " & "{IssueDeppt.gp}='" & cboItemGp.Text & "'"
   End If
   
End Sub

Sub SetPurchaseItemCondition()
   s = ""
   s = "{FinishPurchase.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {FinishPurchase.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yy") & "')"
   
   If cboItem.Text <> "" Then
      s = s & " and " & "{FinishPurchase.ItemName}='" & cboItem.Text & "'"
   End If
   
   If cboSupp.Text <> "" Then
      s = s & " and " & "{FinishPurchase.Supplier}='" & cboSupp.Text & "'"
   End If
End Sub
Sub SetReturnItemCondition()
   s = ""
   s = "{RepairItem.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {RepairItem.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yy") & "')"
   
   If cboItem.Text <> "" Then
      s = s & " and " & "{RepairItem.ItemName}='" & cboItem.Text & "'"
   End If
   
   If cboSupp.Text <> "" Then
      s = s & " and " & "{RepairItem.Supplier}='" & cboSupp.Text & "'"
   End If
End Sub
Private Sub cboItemGp_Click()
   
   If cboItemGp = "" Then Exit Sub
    cboItem.Clear
   If rs.State = 1 Then rs.Close
   rs.Open "select distinct(ItemName) from  ItemCreation where CourseName='" & cboItemGp & "'", CON
   While rs.EOF = False
     cboItem.AddItem rs(0)
     rs.MoveNext
   Wend

End Sub

Private Sub cmdPrint_Click()
   
''   'SetCondition
''   s = ""
''
''   s = "{RawStockSummary.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {RawStockSummary.dates}<=datevalue('" & Format(todate.Value, "MM/dd/yy") & "')"
''
''   If cboItem.Text <> "" Then
''      s = s & " and " & "{RawStockSummary.ItemName}='" & cboItem.Text & "'"
''   End If
''
''
''
''
''
''   cr.Reset
''   cr.ReportFileName = App.Path & "\reports\stocksummary.rpt"
''   cr.Connect = constr
''   cr.ReplaceSelectionFormula s
''   cr.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
''   cr.Formulas(1) = "todate='" & todate.Value & "'"
''   cr.WindowState = crptMaximized
''   cr.WindowShowPrintBtn = True
''   cr.WindowShowPrintSetupBtn = True
''   cr.Action = 1

frmConsumeItemSummary.Show

End Sub

Private Sub Command1_Click()
   
   
   
   
Dim unit As String

If DateValue(FromDate.Value) > DateValue(ToDate.Value) Then
   MsgBox "Invalid Month Selection..", vbCritical
   Exit Sub
End If



Screen.MousePointer = vbHourglass
Dim opening As Long
Dim search As New ADODB.Recordset
Dim save As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
opening = 0

CON.Execute "delete from ConsumeItemStockSummary where len(Name)>0"

If rs.State = 1 Then rs.Close
rs.Open "select * from ItemCreation", CON



If save.State = 1 Then save.Close
save.Open "select * from ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic


If rs.EOF = False Then
'vs.Rows = rs.RecordCount + 1

For i = 1 To rs.RecordCount

unit = rs!unit

Receive = 0
opening = 0
Issue = 0

If rs1.State = 1 Then rs1.Close
rs1.Open "select Opening from itemcreation where itemname='" & rs!itemname & "'", CON
If rs1.EOF = False Then
opening = rs1(0)
End If

If rs1.State = 1 Then rs1.Close
rs1.Open "select sum(Qty) from FinishPurchase where Dates<datevalue('" & FromDate.Value & "') and itemname='" & rs!itemname & "'", CON
If Not IsNull(rs1(0)) Then
opening = opening + rs1(0)
End If

If rs1.State = 1 Then rs1.Close
rs1.Open "select sum(Qty) from IssueDeppt where Dates<datevalue('" & FromDate.Value & "') and itemname='" & rs!itemname & "'", CON
If Not IsNull(rs1(0)) Then
opening = opening - rs1(0)
End If



If rs1.State = 1 Then rs1.Close
rs1.Open "select sum(Qty) from FinishPurchase where (Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "')) and itemname='" & rs!itemname & "'", CON
If Not IsNull(rs1(0)) Then
Receive = rs1(0)
End If

If rs1.State = 1 Then rs1.Close
rs1.Open "select sum(Qty) from IssueDeppt where (Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "')) and itemname='" & rs!itemname & "'", CON
If Not IsNull(rs1(0)) Then
Issue = rs1(0)
End If



save.addNew
save!unit = unit
save!Name = rs!itemname
save!ReceiveStock = Receive
save!Issue = Issue
save!OpenStock = opening
save!ClosingStock = ((opening + Receive) - Issue)
save.Update
rs.MoveNext


opening = 0

Next




End If

Screen.MousePointer = vbDefault
   
   
If MsgBox("Want To View ?", vbQuestion + vbYesNo) = vbYes Then
   
   CR.Reset
   CR.ReportFileName = App.Path & "\DailyConsumption.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula "{ConsumeItemStockSummary.Issue}>0"
   CR.Formulas(0) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
   
End If

End Sub

Private Sub Command10_Click()
   SetReturnItemCondition
   CR.Reset
   CR.ReportFileName = App.Path & "\BalanceRep.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
End Sub

Private Sub Command11_Click()
 SetReturnItemCondition
   CR.Reset
   CR.ReportFileName = App.Path & "\SupplierBalanceRep.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
End Sub

Private Sub Command12_Click()

 SetReturnItemCondition
   CR.Reset
   CR.ReportFileName = App.Path & "\repirSummary.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1

End Sub

Private Sub Command13_Click()

   s = ""
   s = "{NonRetItem.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {NonRetItem.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yy") & "')"
   
   If cboItem.Text <> "" Then
      s = s & " and " & "{NonRetItem.ItemName}='" & cboItem.Text & "'"
   End If
   
   If cboSupp.Text <> "" Then
      s = s & " and " & "{NonRetItem.Supplier}='" & cboSupp.Text & "'"
   End If


   CR.Reset
   CR.ReportFileName = App.Path & "\CollegeWisePending.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
End Sub

Private Sub Command2_Click()
  SetCondition
  CR.Reset
   CR.ReportFileName = App.Path & "\reports\ItemWiseForAll.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
End Sub
Private Sub Command3_Click()
   SetIssueItemCondition
   CR.Reset
   CR.ReportFileName = App.Path & "\reports\CollegeDeptWiseIyemIssue.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
End Sub
Private Sub Command4_Click()
''   SetIssueItemCondition
''   cr.Reset
''   cr.ReportFileName = App.Path & "\DeptWiseAllItemIssue.rpt"
''   cr.Connect = constr
''   cr.ReplaceSelectionFormula s
''   cr.Formulas(0) = "fromdate='" & fromdate.Value & "'"
''   cr.Formulas(1) = "todate='" & todate.Value & "'"
''   cr.WindowState = crptMaximized
''   cr.WindowShowPrintBtn = True
''   cr.WindowShowPrintSetupBtn = True
''   cr.Action = 1

   SetIssueItemCondition
   CR.Reset
   CR.ReportFileName = App.Path & "\reports\BearerWise.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
End Sub
Private Sub Command5_Click()
   SetIssueItemCondition
   CR.Reset
   CR.ReportFileName = App.Path & "\reports\ItemWiseforAllDeptt_Issue.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
End Sub

Private Sub Command6_Click()
Unload Me
End Sub
Private Sub Command7_Click()
   SetPurchaseItemCondition
   CR.Reset
   CR.Connect = constr
   CR.ReportFileName = App.Path & "\reports\ItemWiseSupply.rpt"
   
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
End Sub
Private Sub Command8_Click()
   SetPurchaseItemCondition
   CR.Reset
   CR.ReportFileName = App.Path & "\reports\SupplierWiseItemPurchse.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s
   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
End Sub
Private Sub Command9_Click()
  
   
   s = ""
   s = "{PurchaseReg.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {PurchaseReg.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yy") & "')"
   
   If cboItem.Text <> "" Then
      s = s & " and " & "{PurchaseReg.ItemName}='" & cboItem.Text & "'"
   End If
   
   If cboSupp.Text <> "" Then
      s = s & " and " & "{PurchaseReg.Supplier}='" & cboSupp.Text & "'"
   End If
   
   If cboItemGp.Text <> "" Then
      s = s & " and " & "{PurchaseReg.gp}='" & cboItemGp.Text & "'"
   End If
   
   CR.Reset
   CR.ReportFileName = App.Path & "\Reports\PurchseReg.rpt"
   CR.Connect = constr
   CR.ReplaceSelectionFormula s

   CR.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "todate='" & ToDate.Value & "'"
   CR.WindowState = crptMaximized
   CR.WindowShowPrintBtn = True
   CR.WindowShowPrintSetupBtn = True
   CR.Action = 1
   
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      Unload Me
   End If
End Sub
Private Sub Form_Load()

   'On Error Resume Next
   FromDate.Value = Date
   ToDate.Value = Date
   frmWeatorWiseSale.BackColor = &HDEFEF8
   
   cboCollege.AddItem "ChitraExport"
   cboCollege.ListIndex = 0
   
   If rs.State = 1 Then rs.Close
   rs.Open "select distinct(Name) from  deptt", CON
   While rs.EOF = False
   cboDeptt.AddItem rs(0)
   rs.MoveNext
   Wend
   
       

   If rs.State = 1 Then rs.Close
   rs.Open "select distinct(subledger) from  sledger where gledger='SUNDRY CREDITORS'", CON
   While rs.EOF = False
   cboSupp.AddItem rs(0)
   rs.MoveNext
   Wend
   
   cboItemGp.Clear
   
   If rs.State = 1 Then rs.Close
   rs.Open "select distinct(CourseName) from  ItemCreation", CON
   While rs.EOF = False
   cboItemGp.AddItem rs(0)
   rs.MoveNext
   Wend
   
   
   

End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      ToDate.SetFocus
   End If
End Sub
Private Sub todate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call cmdPrint_Click
End Sub
