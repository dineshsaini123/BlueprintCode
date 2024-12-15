VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPaymentSupply 
   Caption         =   "Party Wise Outst. & T.Supply And Payment"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   18312
   Icon            =   "frmPaymentSupply.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   18312
   Begin VB.CommandButton cmdFilterData 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Filter Data"
      Height          =   645
      Left            =   6525
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   135
      Width           =   1095
   End
   Begin VB.CommandButton CommandReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   645
      Left            =   9960
      Picture         =   "frmPaymentSupply.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CommandPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   645
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   135
      Width           =   1095
   End
   Begin VB.CommandButton Command1_excel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export To Excel"
      Height          =   645
      Left            =   8835
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin Crystal.CrystalReport cr 
      Left            =   11295
      Top             =   360
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker txtFromSale 
      Height          =   315
      Left            =   1380
      TabIndex        =   3
      Top             =   90
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   550
      _Version        =   393216
      Format          =   141164545
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txttoSale 
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   90
      Width           =   1410
      _ExtentX        =   2498
      _ExtentY        =   550
      _Version        =   393216
      Format          =   141164545
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txtFromPayment 
      Height          =   315
      Left            =   1380
      TabIndex        =   7
      Top             =   540
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   550
      _Version        =   393216
      Format          =   141164545
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txttoPayment 
      Height          =   315
      Left            =   3240
      TabIndex        =   8
      Top             =   540
      Width           =   1410
      _ExtentX        =   2498
      _ExtentY        =   550
      _Version        =   393216
      Format          =   141164545
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txt_ason 
      Height          =   315
      Left            =   4815
      TabIndex        =   11
      Top             =   90
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   550
      _Version        =   393216
      Format          =   141164545
      CurrentDate     =   42409
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7590
      Left            =   45
      TabIndex        =   12
      Top             =   1305
      Width           =   18075
      _cx             =   31882
      _cy             =   13388
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   7917545
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPaymentSupply.frx":0BF0
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Filter (From Last year to Current Year)"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1395
      TabIndex        =   14
      Top             =   945
      Width           =   3300
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Range :"
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   540
      Width           =   1350
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   2835
      TabIndex        =   9
      Top             =   540
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Range :"
      Height          =   255
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   1350
   End
   Begin VB.Label lblAson 
      Caption         =   "To"
      Height          =   255
      Left            =   2835
      TabIndex        =   5
      Top             =   90
      Width           =   315
   End
End
Attribute VB_Name = "frmPaymentSupply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFilterData_Click()

Dim con_n As New ADODB.Connection

Screen.MousePointer = vbHourglass

con.Execute "delete from templedger8"

con.Execute "INSERT INTO templedger8 (Balance,drcr,party,billtype,rptid,rptype,setupid,fyear,district,userid,states,Party1)  SELECT op,drcr,subledger,'Opening',1,'ALL',setupid,fyear,ADDRESS3," & UId & ",states,DESCFORINVOICE from sledger group by op,subledger,drcr,setupid,Fyear,ADDRESS3,states,DESCFORINVOICE  HAVING  op <> 0"
con.Execute "INSERT INTO templedger8 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales Bilty No-' + BILTYNO + ',Bundle-' + bundles ,netamount,BAA,SUBLEDGER,fyear,setupid," & UId & ",City,'1','ALL',states,Party,AgentName,scname  from invoiceaQry where  convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger8 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item',BAA,netamount,SUBLEDGER,fyear,setupid," & UId & ",City,'1','ALL',states,Party,AgentName,scname from CREDITAQry where convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
    
con.Execute "INSERT INTO templedger8 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName)" & _
"SELECT  CASHA.INVOICEDATE,'C/M',CASHA.INVOICENO,'Cash Memo',CASHA.NETAMOUNT,CASHA.BAA,CASHA.cashpartyname,CASHA.Fyear," & _
"CASHA.setupid," & UId & ",SLEDGER.ADDRESS3,'1','ALL',SLEDGER.states,SLEDGER.DESCFORINVOICE,CASHA.AgentName " & _
"FROM CASHA INNER JOIN SLEDGER ON CASHA.SUBLEDGER = SLEDGER.SUBLEDGER where convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
    
       
con.Execute "INSERT INTO templedger8 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName)" & _
"SELECT CNF1A.CND,'CN',CNF1A.cnn,'Credit Note ' + desc_ ,0,CNF1A.NA,CNF1A.psld,CNF1A.Fyear," & _
"CNF1A.setupid," & UId & ",SLEDGER.ADDRESS3,'1','ALL',SLEDGER.states,SLEDGER.DESCFORINVOICE,CNF1A.AgentName " & _
"FROM  dbo.CNF1A INNER JOIN SLEDGER ON CNF1A.psld = dbo.SLEDGER.SUBLEDGER where convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)"
    
   
con.Execute "INSERT INTO templedger8 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName)" & _
"SELECT DNFA.DND,'CN',DNFA.Dnn,'Debit Note',DNFA.NA,0,DNFA.psld,DNFA.Fyear," & _
"DNFA.setupid," & UId & ",SLEDGER.ADDRESS3,'1','ALL',SLEDGER.states,SLEDGER.DESCFORINVOICE,DNFA.AgentName " & _
"FROM DNFA INNER JOIN SLEDGER ON DNFA.psld = SLEDGER.SUBLEDGER where convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & txt_ason.value & "',103)"
    
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
    
con.Execute "INSERT INTO templedger8 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) " & _
" SELECT a.Dates,'J',a.RecNo, a.Particullar, a.Dr, a.Cr,a.PartyName,a.fyear,a.setupid," & UId & "," & _
" b.ADDRESS3,'1','ALL',b.states,b.DESCFORINVOICE,b.repname1 FROM ReceiveIssueParty as a INNER JOIN " & _
" SLEDGER as b ON a.PartyName = b.SUBLEDGER where convert(smalldatetime,DATEs,103)<=convert(smalldatetime,'" & txt_ason.value & "',103) and a.firm='chitra'  order by dates,recno"




con.Execute "delete from tmpPayment"

con.Execute "exec SpTotalSupply '" & txtFromSale.value & "','" & txttoSale.value & "'"

'------------------------------------------------------------------------------------------
Set con_n = New ADODB.Connection

If RS.State = 1 Then RS.close
RS.Open "select DataBase,NotCreated from turnOverDis where Current_Next='next'", CCON
If RS.EOF = False Then
   
   If RS!NotCreated = "n" Then
      con_n.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & last_dbase & "; UID=" & sql_user & "; PWD=" & sql_pass
      con_n.Open
   Else
      con_n.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & next_dbase & "; UID=" & sql_user & "; PWD=" & sql_pass
      con_n.Open
   End If

End If


'------------------------------------------------------------------------------------------

If RS.State = 1 Then RS.close
RS.Open "select Code,Party,City,states,NetAmount,agentname,PartyTerms,subledger,INVOICEDATE from Yearly_PartyNetSupply where (INVOICEDATE>=convert(smalldatetime,'" + Trim(txtFromSale.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(txttoSale.value) + "',103))", con_n
While RS.EOF = False

con.Execute "insert into TmpYearly_PartyNetSupply(Code,Party,City,states,NetAmount,agentname,PartyTerms,subledger,INVOICEDATE) " & _
" values('" & RS!Code & "','" & RS!party & "','" & RS!city & "','" & RS!states & "','" & RS!netamount & "','" & RS!agentname & "','" & RS!PartyTerms & "','" & RS!subledger & "','" & Format(RS!invoiceDate, "mm/dd/yyyy") & "')"

RS.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "select VoucherType , VoucherDate, Genledger, SubLedger, amount, DebitorCredit,DESCRIPTION from VOUCHERS where (GenLedger='SUNDRY DEBTORS' and SUBLEDGER not like '%IMPREST A/C%') and VoucherDate>=convert(smalldatetime,'" + Trim(txtFromPayment.value) + "',103) and VoucherDate<=convert(smalldatetime,'" + Trim(txttoPayment.value) + "',103)", con_n
While RS.EOF = False

con.Execute "insert into tmpPayment(VoucherType , VoucherDate, Genledger, SubLedger, amount, DebitorCredit,DESCRIPTION) " & _
" values('" & RS!vouchertype & "','" & Format(RS!voucherDATE, "mm/dd/yyyy") & "','" & RS!Genledger & "','" & RS!subledger & "','" & RS!amount & "','" & RS!DebitorCredit & "','" & RS!DESCRIPTION & "')"

RS.MoveNext
Wend

'---------------------------------------------------------------------------------------------

con.Execute "exec SpTotalPayment '" & txtFromPayment.value & "','" & txttoPayment.value & "'"



DoEvents
DoEvents
DoEvents
DoEvents
DoEvents

Screen.MousePointer = vbDefault



End Sub
Private Sub Command1_excel_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim str_ As String




If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double



row_ = 1
col_ = 1

xl.Columns("A:H").ColumnWidth = 12
J = 2


For I = 0 To vs.rows - 1
    For J = 0 To vs.Cols - 1
      
        xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
       
        col_ = col_ + 1
    Next
    row_ = row_ + 1
    col_ = 1
Next

MsgBox "Task Completed....", vbInformation

End Sub
Private Sub CommandPrint_Click()


Screen.MousePointer = vbHourglass


con.Execute "UPDATE a SET a.balance = b.op  FROM templedger8 AS a " & _
"INNER JOIN SLEDGER AS b ON (a.party = b.subledger)"

con.Execute "UPDATE a SET a.drcr = b.drcr  FROM templedger8 AS a " & _
" INNER JOIN SLEDGER AS b ON (a.party = b.subledger)"


con.Execute "UPDATE a SET a.TSale = b.NetAmount  FROM templedger8 AS a " & _
"LEFT JOIN Yearly_PartyWiseSupplySummary as b ON (a.party = b.subledger)"


con.Execute "UPDATE templedger8 SET balance = (balance*-1)  where drcr='CR'"
  

con.Execute "UPDATE a SET a.TPay = b.NetAmount  FROM templedger8 AS a " & _
"LEFT JOIN Yearly_PartyWisePaymentSummary as b ON (a.party = b.subledger)"


If MsgBox("Want to view ?", vbQuestion + vbYesNo, "Message") = vbNo Then
   Screen.MousePointer = vbDefault
   Exit Sub
End If

DoEvents
DoEvents
DoEvents

'============================================================================
Set rs1 = New ADODB.Recordset
DoEvents
DoEvents
''rs1.Open "SELECT substring(Party,1,5) as code,sum(CollectedAmt) as CollectedAmt FROM ApprovalCollectedAmt group by substring(Party,1,5)", con, adOpenDynamic, adLockOptimistic
rs1.Open "SELECT substring(Party,1,5) as code,sum(CollectedAmt) as CollectedAmt,sum(collectedAmt_net) as collectedAmt_net,Net_Gross FROM ApprovalCollectedAmt group by substring(Party,1,5),Net_Gross", con, adOpenDynamic, adLockOptimistic
DoEvents
'============================================================================
vs.rows = 2
vs.Cols = 12
vs.FormatString = "Code|Party|PartyTerms|City|State|OutStanding|RepName1|RepName2|RepName3|TSale|TPayment|AdjAmt.|collectedAmt."

op_ = 0
dr_ = 0
cr_ = 0
outs_ = 0

I = 1
If RS.State = 1 Then RS.close
RS.Open "SELECT substring(a.Party,1,5) as Code,b.DESCFORINVOICE,a.District as City,a.states,a.Balance as OP,sum(a.Dr) as Dr,sum(a.Cr) as Cr,a.TSale,a.TPay,b.RepName1,b.RepName2,b.RepName3,b.PartyRemarks from templedger8 as a inner join sledger as b on (a.Party = b.subledger) group by a.Party,b.DESCFORINVOICE,a.District,a.states,a.Balance,a.TSale,a.TPay,b.RepName1,b.RepName2,b.RepName3,b.PartyRemarks", con
While RS.EOF = False
 
 DoEvents
 
 If Not IsNull(RS(4)) Then
 op_ = RS(4)
 Else
 op_ = 0
 End If
 
 If Not IsNull(RS(5)) Then
 dr_ = RS(5)
 Else
 dr_ = 0
 End If
 
 If Not IsNull(RS(6)) Then
 cr_ = (RS(6) * -1)
 Else
 cr_ = 0
 End If
 
 outs = op_ + dr_ + cr_
 
 
 
 vs.TextMatrix(I, 0) = RS(0)
 vs.TextMatrix(I, 1) = RS(1)
 vs.TextMatrix(I, 2) = RS.Fields("PartyRemarks").value & ""
 
 vs.TextMatrix(I, 3) = RS(2)
 vs.TextMatrix(I, 4) = RS(3)
 vs.TextMatrix(I, 5) = Round(outs, 2)
 
 vs.TextMatrix(I, 6) = RS.Fields("RepName1").value
 vs.TextMatrix(I, 7) = RS.Fields("RepName2").value
 vs.TextMatrix(I, 8) = RS.Fields("RepName3").value
 
 vs.TextMatrix(I, 9) = IIf(IsNull(RS!TSale), 0, RS!TSale)
 vs.TextMatrix(I, 10) = IIf(IsNull(RS!TPay), 0, RS!TPay)
 
 If RS(0) = "G2519" Then
 'MsgBox "s"
 End If
 
 On Error Resume Next
 rs1.MoveFirst
 If rs1.EOF = False Then
    
    rs1.Find "code='" & RS(0) & "'"
    If rs1.EOF = False Then
       If rs1!Net_Gross = "Gross" Then
          vs.TextMatrix(I, 11) = IIf(IsNull(rs1(1)), 0, Round(rs1(1), 2))
       Else
          vs.TextMatrix(I, 11) = IIf(IsNull(rs1(2)), 0, Round(rs1(2), 2))
       End If
       vs.TextMatrix(I, 11) = Round(vs.TextMatrix(I, 11), 2)
    End If
 Else
    vs.TextMatrix(I, 11) = IIf(vs.TextMatrix(I, 11) = "", 0, vs.TextMatrix(I, 11))
 End If
 
  
 vs.TextMatrix(I, 12) = (IIf(vs.TextMatrix(I, 5) = "", 0, vs.TextMatrix(I, 5)) - IIf(vs.TextMatrix(I, 11) = "", 0, vs.TextMatrix(I, 11)))
 vs.TextMatrix(I, 12) = Round(vs.TextMatrix(I, 12), 2)
 
 DoEvents
 DoEvents
 DoEvents
 
 vs.rows = vs.rows + 1
 I = I + 1
 RS.MoveNext
Wend

vs.ColWidth(0) = 700
vs.ColWidth(1) = 2000
vs.ColWidth(2) = 1700
vs.ColWidth(3) = 1700
vs.ColWidth(4) = 1200
vs.ColWidth(5) = 1200
vs.ColWidth(6) = 1500
vs.ColWidth(7) = 1500
vs.ColWidth(8) = 1200
vs.ColWidth(9) = 1100
vs.ColWidth(10) = 1200
vs.ColWidth(11) = 1200


Screen.MousePointer = vbDefault



End Sub
Private Sub CommandReturn_Click()
 Unload Me
End Sub
Private Sub Form_Load()

Me.Width = 18500
Me.Height = 10000

txtFromPayment.value = Format(Date, "dd/MM/yyyy")
txttoPayment.value = Format(Date, "dd/MM/yyyy")

txtFromSale.value = Format(Date, "dd/MM/yyyy")
txttoSale.value = Format(Date, "dd/MM/yyyy")

txt_ason.value = Format(Date, "dd/MM/yyyy")

BackColorFrom Me
End Sub
