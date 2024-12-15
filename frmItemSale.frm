VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmItemSale 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3192
   ClientLeft      =   3168
   ClientTop       =   3108
   ClientWidth     =   5712
   Icon            =   "frmItemSale.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   5712
   Begin VB.ListBox complist 
      Height          =   696
      Left            =   1590
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   1500
      Width           =   3795
   End
   Begin Crystal.CrystalReport cr 
      Left            =   60
      Top             =   2250
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker fromdate 
      Height          =   300
      Left            =   1575
      TabIndex        =   1
      Top             =   705
      Width           =   1425
      _ExtentX        =   2519
      _ExtentY        =   529
      _Version        =   393216
      Format          =   156631041
      CurrentDate     =   38915
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   300
      Left            =   1575
      TabIndex        =   2
      Top             =   1125
      Width           =   1425
      _ExtentX        =   2519
      _ExtentY        =   529
      _Version        =   393216
      Format          =   156631041
      CurrentDate     =   38915
   End
   Begin VB.ComboBox cboItem 
      Height          =   315
      Left            =   1575
      TabIndex        =   0
      Top             =   240
      Width           =   3750
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   390
      Left            =   2895
      TabIndex        =   4
      Top             =   2475
      Width           =   1770
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   390
      Left            =   1140
      TabIndex        =   3
      Top             =   2475
      Width           =   1770
   End
   Begin VB.Label lblcname 
      Caption         =   "Company Name"
      Height          =   435
      Left            =   405
      TabIndex        =   9
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "To Date"
      Height          =   255
      Left            =   405
      TabIndex        =   7
      Top             =   1140
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
      Height          =   255
      Left            =   405
      TabIndex        =   6
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Item Name"
      Height          =   255
      Left            =   405
      TabIndex        =   5
      Top             =   255
      Width           =   1215
   End
End
Attribute VB_Name = "frmItemSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rrs As New ADODB.Recordset

Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub cmdexit_Click()
   Unload Me
End Sub
Sub updatePartyReceipt()

Dim RS As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim party As New ADODB.Recordset
Dim invAmt As Double
Dim recAmt, op As Double
Dim strcid As String
recAmt = 0
invAmt = 0
op = 0

For I = 0 To complist.ListCount - 1
If complist.Selected(I) = True Then
strcid = Val(Left(complist.List(I), 2))
''"SUNDRY DEBTORS"
con.Execute "update invoicea set RecAmt=0,updatedby='" & main.UserName & "',updatedon=CONVERT(smalldatetime,'" & CDate(Now) & "',103) where fyear='" & main.session & "' and setupid=" & strcid & " and RecAmt>0"
con.Execute "update invoicea set lexp3rate=0,updatedby='" & main.UserName & "',updatedon=CONVERT(smalldatetime,'" & Now & "',103) where fyear='" & main.session & "' and setupid=" & strcid & " and lexp3rate>0"
con.Execute "update sledger set tempop=0,updatedby='" & main.UserName & "',updatedon=CONVERT(smalldatetime,'" & Now & "',103) where fyear='" & main.session & "' and setupid=" & strcid & ""

If RS.State = 1 Then RS.close
RS.Open "select SUBLEDGER from invoicea where fyear='" & main.session & "' and setupid=" & strcid & IIf(cboItem.Text <> "", " and district='" & cboItem.Text & "'", "") & " group by SUBLEDGER", con
    While RS.EOF = False

        If rss.State = 1 Then rss.close
        rss.Open "select NETAMOUNT,INVOICENO from invoicea where   fyear='" & main.session & "' and setupid=" & strcid & " and SubLedger='" & RS.Fields(0).value & "' order by INVOICENO", con
        If rss.EOF = False Then
           'R CON.Execute "update invoicea set lexp3rate=" & op & ",updatedby='" & main.username & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  " & stringyear & " and INVOICENO=" & rss.Fields(1).Value & ""
            While rss.EOF = False
                        'invAmt = myround(rss.Fields(0).Value, 0)
                        'If invAmt <= recAmt Then
                        '    CON.Execute "update invoicea set lexp3rate=" & op & ", RecAmt=" & invAmt & ",updatedby='" & main.username & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  " & stringyear & " and INVOICENO=" & rss.Fields(1).Value & ""
                           If rs1.State = 1 Then rs1.close
                           rs1.Open "select sum(amount) as recamt from vouchers where  fyear='" & main.session & "' and setupid=" & strcid & " and SubLedger='" & RS.Fields(0).value & "' AND DebitorCredit='C' and convert(int,cbnd)=" & rss!invoiceNo, con
                           If Not IsNull(rs1(0)) Then
                               recAmt = rs1(0)
                           Else
                               recAmt = 0
                           End If
                        con.Execute "update invoicea set RecAmt=" & recAmt & ",updatedby='" & main.UserName & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  fyear='" & main.session & "' and setupid=" & strcid & " and INVOICENO=" & rss.Fields(1).value & ""
                        '    recAmt = recAmt - invAmt
                        'ElseIf invAmt >= recAmt Then
                        '    CON.Execute "update invoicea set lexp3rate=" & op & ", RecAmt=" & recAmt & ",updatedby='" & main.username & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  " & stringyear & " and INVOICENO=" & rss.Fields(1).Value & ""
                        '    recAmt = 0
                        '    op = 0
                        'End If
                    rss.MoveNext
            Wend
      End If
        
        If party.State = 1 Then party.close
        party.Open "select YEAROPENING,SUBLEDGER from sledger where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "'", con
        If party.EOF = False Then
            op = Val(party.Fields(0).value & "")
        Else
            op = 0
        End If
        
        If rs1.State = 1 Then rs1.close
        'rs1.Open "select sum(Amount) from vouchers where " & stringyear & " and SubLedger='" & rs.Fields(0).Value & "' AND DebitorCredit='C' ", CON
        rs1.Open "select sum(netAmount),sum(recamt) from invoicea where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" & FromDate.value & "',103)", con
        If Not IsNull(rs1(0)) Then
            op = op + rs1(0)
        Else
            op = op
        End If
        
        If Not IsNull(rs1(1)) Then
            recAmt = rs1(1)
        Else
            recAmt = 0
        End If
        
        
        
        If rs1.State = 1 Then rs1.close
        'rs1.Open "select sum(Amount) from vouchers where " & stringyear & " and SubLedger='" & rs.Fields(0).Value & "' AND DebitorCredit='C' ", CON
        rs1.Open "select sum(Amount) from vouchers where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "' AND DebitorCredit='C' and convert(int,cbnd)=0", con
        If Not IsNull(rs1(0)) Then
            recAmt = recAmt + rs1(0)
        Else
            recAmt = recAmt
        End If

        con.Execute "update sledger set tempOp=" & (op) & ",TEMPRECOP=" & recAmt & ",updatedby='" & main.UserName & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "'"


RS.MoveNext
Wend
End If
Next

End Sub


Sub updatePartyReceiptagentwise()

Dim RS As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim party As New ADODB.Recordset
Dim invAmt As Double
Dim recAmt, op As Double
Dim strcid As String
recAmt = 0
invAmt = 0
op = 0

For I = 0 To complist.ListCount - 1
If complist.Selected(I) = True Then
strcid = Val(Left(complist.List(I), 2))
''"SUNDRY DEBTORS"
con.Execute "update invoicea set RecAmt=0,updatedby='" & main.UserName & "',updatedon=CONVERT(smalldatetime,'" & CDate(Now) & "',103) where fyear='" & main.session & "' and setupid=" & strcid & " and RecAmt>0"
con.Execute "update invoicea set lexp3rate=0,updatedby='" & main.UserName & "',updatedon=CONVERT(smalldatetime,'" & Now & "',103) where fyear='" & main.session & "' and setupid=" & strcid & " and lexp3rate>0"
con.Execute "update sledger set tempop=0,updatedby='" & main.UserName & "',updatedon=CONVERT(smalldatetime,'" & Now & "',103) where fyear='" & main.session & "' and setupid=" & strcid & ""

If RS.State = 1 Then RS.close
RS.Open "select SUBLEDGER from invoicea where fyear='" & main.session & "' and setupid=" & strcid & IIf(cboItem.Text <> "", " and agentname='" & cboItem.Text & "'", "") & " group by SUBLEDGER", con
    While RS.EOF = False

        If rss.State = 1 Then rss.close
        rss.Open "select NETAMOUNT,INVOICENO from invoicea where   fyear='" & main.session & "' and setupid=" & strcid & " and SubLedger='" & RS.Fields(0).value & "' order by INVOICENO", con
        If rss.EOF = False Then
           'R CON.Execute "update invoicea set lexp3rate=" & op & ",updatedby='" & main.username & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  " & stringyear & " and INVOICENO=" & rss.Fields(1).Value & ""
            While rss.EOF = False
                        'invAmt = myround(rss.Fields(0).Value, 0)
                        'If invAmt <= recAmt Then
                        '    CON.Execute "update invoicea set lexp3rate=" & op & ", RecAmt=" & invAmt & ",updatedby='" & main.username & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  " & stringyear & " and INVOICENO=" & rss.Fields(1).Value & ""
                           If rs1.State = 1 Then rs1.close
                           rs1.Open "select sum(amount) as recamt from vouchers where  fyear='" & main.session & "' and setupid=" & strcid & " and SubLedger='" & RS.Fields(0).value & "' AND DebitorCredit='C' and convert(int,cbnd)=" & rss!invoiceNo, con
                           If Not IsNull(rs1(0)) Then
                               recAmt = rs1(0)
                           Else
                               recAmt = 0
                           End If
                        con.Execute "update invoicea set RecAmt=" & recAmt & ",updatedby='" & main.UserName & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  fyear='" & main.session & "' and setupid=" & strcid & " and INVOICENO=" & rss.Fields(1).value & ""
                        '    recAmt = recAmt - invAmt
                        'ElseIf invAmt >= recAmt Then
                        '    CON.Execute "update invoicea set lexp3rate=" & op & ", RecAmt=" & recAmt & ",updatedby='" & main.username & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  " & stringyear & " and INVOICENO=" & rss.Fields(1).Value & ""
                        '    recAmt = 0
                        '    op = 0
                        'End If
                    rss.MoveNext
            Wend
      End If
        
        If party.State = 1 Then party.close
        party.Open "select YEAROPENING,SUBLEDGER from sledger where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "'", con
        If party.EOF = False Then
            op = Val(party.Fields(0).value & "")
        Else
            op = 0
        End If
        
        If rs1.State = 1 Then rs1.close
        'rs1.Open "select sum(Amount) from vouchers where " & stringyear & " and SubLedger='" & rs.Fields(0).Value & "' AND DebitorCredit='C' ", CON
        rs1.Open "select sum(netAmount),sum(recamt) from invoicea where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" & FromDate.value & "',103)", con
        If Not IsNull(rs1(0)) Then
            op = op + rs1(0)
        Else
            op = op
        End If
        
        If Not IsNull(rs1(1)) Then
            recAmt = rs1(1)
        Else
            recAmt = 0
        End If
        
        
        
        If rs1.State = 1 Then rs1.close
        'rs1.Open "select sum(Amount) from vouchers where " & stringyear & " and SubLedger='" & rs.Fields(0).Value & "' AND DebitorCredit='C' ", CON
        rs1.Open "select sum(Amount) from vouchers where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "' AND DebitorCredit='C' and convert(int,cbnd)=0", con
        If Not IsNull(rs1(0)) Then
            recAmt = recAmt + rs1(0)
        Else
            recAmt = recAmt
        End If

        con.Execute "update sledger set tempOp=" & (op) & ",TEMPRECOP=" & recAmt & ",updatedby='" & main.UserName & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "'"


RS.MoveNext
Wend
End If
Next

End Sub




Sub updatePartypayment()

Dim RS As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim party As New ADODB.Recordset
Dim invAmt As Double
Dim recAmt, op As Double
Dim strcid As String
recAmt = 0
invAmt = 0
op = 0

For I = 0 To complist.ListCount - 1
If complist.Selected(I) = True Then
strcid = Val(Left(complist.List(I), 2))
''"SUNDRY DEBTORS"
con.Execute "update purchasea set RecAmt=0,updatedby='" & main.UserName & "',updatedon=CONVERT(smalldatetime,'" & CDate(Now) & "',103) where fyear='" & main.session & "' and setupid=" & strcid & " and RecAmt>0"
con.Execute "update purchasea set lexp3rate=0,updatedby='" & main.UserName & "',updatedon=CONVERT(smalldatetime,'" & Now & "',103) where fyear='" & main.session & "' and setupid=" & strcid & " and lexp3rate>0"
con.Execute "update sledger set tempop=0,updatedby='" & main.UserName & "',updatedon=CONVERT(smalldatetime,'" & Now & "',103) where fyear='" & main.session & "' and setupid=" & strcid & ""

If RS.State = 1 Then RS.close
RS.Open "select SUBLEDGER from purchasea where fyear='" & main.session & "' and setupid=" & strcid & IIf(cboItem.Text <> "", " and district='" & cboItem.Text & "'", "") & " group by SUBLEDGER", con
    While RS.EOF = False

        If rss.State = 1 Then rss.close
        rss.Open "select NETAMOUNT,INVOICENO from purchasea where   fyear='" & main.session & "' and setupid=" & strcid & " and SubLedger='" & RS.Fields(0).value & "' order by INVOICENO", con
        If rss.EOF = False Then
           'R CON.Execute "update invoicea set lexp3rate=" & op & ",updatedby='" & main.username & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  " & stringyear & " and INVOICENO=" & rss.Fields(1).Value & ""
            While rss.EOF = False
                        'invAmt = myround(rss.Fields(0).Value, 0)
                        'If invAmt <= recAmt Then
                        '    CON.Execute "update invoicea set lexp3rate=" & op & ", RecAmt=" & invAmt & ",updatedby='" & main.username & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  " & stringyear & " and INVOICENO=" & rss.Fields(1).Value & ""
                           If rs1.State = 1 Then rs1.close
                           rs1.Open "select sum(amount) as recamt from vouchers where  fyear='" & main.session & "' and setupid=" & strcid & " and SubLedger='" & RS.Fields(0).value & "' AND DebitorCredit='D' and convert(int,cbnd)=" & rss!invoiceNo, con
                           If Not IsNull(rs1(0)) Then
                               recAmt = rs1(0)
                           Else
                               recAmt = 0
                           End If
                        con.Execute "update purchasea set RecAmt=" & recAmt & ",updatedby='" & main.UserName & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  fyear='" & main.session & "' and setupid=" & strcid & " and INVOICENO=" & rss.Fields(1).value & ""
                        '    recAmt = recAmt - invAmt
                        'ElseIf invAmt >= recAmt Then
                        '    CON.Execute "update invoicea set lexp3rate=" & op & ", RecAmt=" & recAmt & ",updatedby='" & main.username & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  " & stringyear & " and INVOICENO=" & rss.Fields(1).Value & ""
                        '    recAmt = 0
                        '    op = 0
                        'End If
                    rss.MoveNext
            Wend
      End If
        
        If party.State = 1 Then party.close
        party.Open "select YEAROPENING,SUBLEDGER from sledger where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "'", con
        If party.EOF = False Then
            op = Val(party.Fields(0).value & "")
        Else
            op = 0
        End If
        
        If rs1.State = 1 Then rs1.close
        'rs1.Open "select sum(Amount) from vouchers where " & stringyear & " and SubLedger='" & rs.Fields(0).Value & "' AND DebitorCredit='C' ", CON
        rs1.Open "select sum(netAmount),sum(recamt) from invoicea where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" & FromDate.value & "',103)", con
        If Not IsNull(rs1(0)) Then
            op = op + rs1(0)
        Else
            op = op
        End If
        
        If Not IsNull(rs1(1)) Then
            recAmt = rs1(1)
        Else
            recAmt = 0
        End If
        
        
        
        If rs1.State = 1 Then rs1.close
        'rs1.Open "select sum(Amount) from vouchers where " & stringyear & " and SubLedger='" & rs.Fields(0).Value & "' AND DebitorCredit='C' ", CON
        rs1.Open "select sum(Amount) from vouchers where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "' AND DebitorCredit='D' and convert(int,cbnd)=0", con
        If Not IsNull(rs1(0)) Then
            recAmt = recAmt + rs1(0)
        Else
            recAmt = recAmt
        End If

        con.Execute "update sledger set tempOp=" & (op) & ",TEMPRECOP=" & recAmt & ",updatedby='" & main.UserName & "',updatedon=convert(smalldatetime,'" & Now & "',103) where  fyear='" & main.session & "' and setupid=" & strcid & "  and SubLedger='" & RS.Fields(0).value & "'"


RS.MoveNext
Wend
End If
Next

End Sub
Private Sub cmdPrint_Click()
    Dim strcid As String
    
    DSNNew
    
    For I = 0 To complist.ListCount - 1
    If complist.Selected(I) = True Then
    strcid = strcid & Val(Left(arycname(I), 2)) & ","
    End If
    
    
    
    Next
    strcid = Mid(strcid, 1, Len(strcid) - 1)
    
 If ss1 = 1 Then
    
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = rptPath & "\ItemWiseSales.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If cboItem.Text <> "" Then
      cr.SelectionFormula = "{invoiceb.fyear}='" & main.session & "' and {invoiceb.setupid} in [" & strcid & "] and {books.BOOKNAME} = '" & cboItem.Text & "' and {invoiceb.INVOICEDATE}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {invoiceb.INVOICEDATE}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"
    Else
      cr.SelectionFormula = "{invoiceb.fyear}='" & main.session & "' and {invoiceb.setupid} in [" & strcid & "] and {invoiceb.INVOICEDATE}>=datevalue('" & Format(FromDate.value, "MM/dd/yy") & "') and {invoiceb.INVOICEDATE}<=datevalue('" & Format(toDate.value, "MM/dd/yy") & "')"
    End If
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
    
 ElseIf ss1 = 2 Then
  Screen.MousePointer = vbHourglass
    
    updatePartyReceiptagentwise
    
    Screen.MousePointer = vbDefault
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = rptPath & "\ItemWiseAgent.rpt"
    If cboItem.Text <> "" Then
'       cr.SelectionFormula = "{invoicea.fyear}='" & main.session & "' and {invoicea.setupid} in [" & strcid & "] and {INVOICEA.AgentName} = '" & cboItem.Text & "' and {invoicea.INVOICEDATE}>=datevalue('" & fromdate.Value & "') and {invoicea.INVOICEDATE}<=datevalue('" & todate.Value & "')"
'    Else
'      cr.SelectionFormula = "{invoicea.fyear}='" & main.session & "' and  {invoicea.setupid} in [" & strcid & "] and {INVOICEa.INVOICEDATE}>=datevalue('" & fromdate.Value & "') and {invoicea.INVOICEDATE}<=datevalue('" & todate.Value & "')"
    
         cr.ReplaceSelectionFormula "{invoicea.fyear}='" & main.session & "' and {invoicea.setupid} in [" & strcid & "] and {invoicea.agentname} = '" & cboItem.Text & "'  and  {invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"
   Else
        cr.ReplaceSelectionFormula "{invoicea.fyear}='" & main.session & "' and {invoicea.setupid} in [" & strcid & "] and {invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"
 
    End If
    cr.Formulas(0) = "fromdate='" & FromDate.value & "'"
    cr.Formulas(1) = "todate='" & toDate.value & "'"

    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
    
 ElseIf ss1 = 3 Then
    
    
    Screen.MousePointer = vbHourglass
    
    updatePartyReceipt
    
    Screen.MousePointer = vbDefault
    
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = rptPath & "\PartyWiseStm.rpt"
    If cboItem.Text <> "" Then
'       'CR.SelectionFormula = "(({invoicea.netAmount}-{invoicea.recamt})>0) and {invoicea.district} = '" & cboItem.Text & "' and  ({invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "'))"
        'cr.SelectionFormula = "{invoicea.district} = '" & cboItem.Text & "'  and  {invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
        cr.ReplaceSelectionFormula "{invoicea.fyear}='" & main.session & "' and {invoicea.setupid} in [" & strcid & "] and {invoicea.district} = '" & cboItem.Text & "'  and  {invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"
   Else
        cr.ReplaceSelectionFormula "{invoicea.fyear}='" & main.session & "' and {invoicea.setupid} in [" & strcid & "] and {invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"

'         'CR.SelectionFormula = "(({invoicea.netAmount}-{invoicea.recamt})>0) and  ({invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "'))"
'         cr.SelectionFormula = "({invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "'))"
    End If
    cr.Formulas(0) = "fromdate='" & FromDate.value & "'"
    cr.Formulas(1) = "todate='" & toDate.value & "'"

    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
 
 ElseIf ss1 = 4 Then
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = rptPath & "\QualityWiseItemList.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
        
    If cboItem.Text <> "" Then
        cr.SelectionFormula = "{invoicea.fyear}='" & main.session & "' and {invoicea.setupid} in [" & strcid & "] and {books.QUALITY} = '" & cboItem.Text & "' and {invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.value, "MM/dd/yy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(toDate.value, "MM/dd/yy") & "')"
    Else
        cr.SelectionFormula = "{invoicea.fyear}='" & main.session & "' and {invoicea.setupid} in [" & strcid & "] and {invoicea.INVOICEDATE}>=datevalue('" & FromDate.value & "') and {invoicea.INVOICEDATE}<=datevalue('" & toDate.value & "')"
    End If
    cr.Formulas(0) = "fromdate='" & FromDate.value & "'"
    cr.Formulas(1) = "todate='" & toDate.value & "'"
    
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
 ElseIf ss1 = 5 Then
    
    
    Screen.MousePointer = vbHourglass
    
    updatePartypayment
    
    Screen.MousePointer = vbDefault
    
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = rptPath & "\PartyWisepurchaseStm.rpt"
    If cboItem.Text <> "" Then
'       'CR.SelectionFormula = "(({invoicea.netAmount}-{invoicea.recamt})>0) and {invoicea.district} = '" & cboItem.Text & "' and  ({invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "'))"
        'cr.SelectionFormula = "{invoicea.district} = '" & cboItem.Text & "'  and  {invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
        cr.ReplaceSelectionFormula "{invoicea.fyear}='" & main.session & "' and {invoicea.setupid} in [" & strcid & "] and {invoicea.district} = '" & cboItem.Text & "'  and  {invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"
   Else
        cr.ReplaceSelectionFormula "{invoicea.fyear}='" & main.session & "' and {invoicea.setupid} in [" & strcid & "] and {invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(toDate.value, "MM/dd/yyyy") & "')"

'         'CR.SelectionFormula = "(({invoicea.netAmount}-{invoicea.recamt})>0) and  ({invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "'))"
'         cr.SelectionFormula = "({invoicea.INVOICEDATE}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {invoicea.INVOICEDATE}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "'))"
    End If
    cr.Formulas(0) = "fromdate='" & FromDate.value & "'"
    cr.Formulas(1) = "todate='" & toDate.value & "'"

    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
 End If
    
End Sub

Private Sub complist_ItemCheck(Item As Integer)
If VBA.Left(complist.List(Item), InStr(1, complist.List(Item), " ")) = main.setupid Then
complist.Selected(Item) = True
Else

End If
End Sub
Private Sub Form_Load()
complist.Visible = main.blnviewallcomp = True
lblcname.Visible = main.blnviewallcomp = True

For I = 0 To UBound(arycname)
complist.AddItem arycname(I)
If Val(Left(arycname(I), 2)) = main.setupid Then
complist.Selected(I) = True
End If

Next

Dim RS As New ADODB.Recordset
Set RS = New ADODB.Recordset

RS.Open "Select * from setup where " & stringyear & "", con, adOpenStatic, adLockReadOnly, adCmdText
CNSetup
FromDate.value = RS!yarfrom
toDate.value = RS!yarto
RS.close

If ss1 = 1 Then

Label1.Caption = "Item Name"
If rrs.State = 1 Then rrs.close
rrs.Open "select distinct(BOOKNAME) from books where " & stringyear & " ", con
While rrs.EOF = False
    cboItem.AddItem rrs(0)
    rrs.MoveNext
Wend
ElseIf ss1 = 2 Then

Label1.Caption = "Agent Name"
If rrs.State = 1 Then rrs.close
rrs.Open "select distinct(AgentName) from INVOICEA  where " & stringyear & "", con
While rrs.EOF = False
    cboItem.AddItem rrs(0)
    rrs.MoveNext
Wend

ElseIf ss1 = 3 Then
Label1.Caption = "Station"
If rrs.State = 1 Then rrs.close
rrs.Open "select distinct(district) from INVOICEA where  " & stringyear & "", con
While rrs.EOF = False
    cboItem.AddItem rrs(0) & ""
    rrs.MoveNext
Wend

ElseIf ss1 = 4 Then
Label1.Caption = "Quality"
If rrs.State = 1 Then rrs.close
rrs.Open "select distinct(QUALITY) from books where " & stringyear & " ", con
While rrs.EOF = False
   If Not IsNull(rrs(0)) Then
    cboItem.AddItem rrs(0)
   End If
    rrs.MoveNext
Wend
ElseIf ss1 = 5 Then
Label1.Caption = "Station"
If rrs.State = 1 Then rrs.close
rrs.Open "select distinct(district) from purchaseA where  " & stringyear & "", con
While rrs.EOF = False
    cboItem.AddItem rrs(0) & ""
    rrs.MoveNext
Wend
End If
End Sub

Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Todate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub
