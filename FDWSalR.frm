VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DWSalesReturn 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   6015
   ClientTop       =   2220
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox AgCombo 
      Height          =   315
      Left            =   2460
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3885
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   4350
      TabIndex        =   5
      Top             =   2970
      Width           =   1455
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   2580
      TabIndex        =   4
      Top             =   2970
      Width           =   1545
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   2460
      TabIndex        =   2
      Top             =   1590
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox Combosldistrictcode 
      Height          =   315
      Left            =   2460
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3885
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   4350
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "District Name"
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   1140
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   315
      Left            =   3840
      TabIndex        =   8
      Top             =   1620
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From The Date"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   1620
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Agent Name"
      Height          =   195
      Left            =   660
      TabIndex        =   6
      Top             =   630
      Width           =   885
   End
End
Attribute VB_Name = "DWSalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Dim RS As Recordset
Private Sub COMBOGENLEDGER_Change()
'    COMBOGENLEDGER = UCase(COMBOGENLEDGER)
'    If rs.State = 1 Then
'        rs.Close
'    End If
  '  rs.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
  '  Combosubledger.Clear
  '  If Not rs.BOF Then
   '     Do While Not rs.EOF
    '        Combosubledger.AddItem Trim(rs!subledger)
     '       If Not rs.EOF Then
      '          rs.MoveNext
       '     End If
       ' Loop
  '  End If
  '  rs.Close
    
End Sub
Private Sub COMBOGENLEDGER_LostFocus()
COMBOGENLEDGER = UCase(COMBOGENLEDGER)
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        RS.Open "select * from gledger where  " & stringyear & " and slf=0", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.BOF Then
                COMBOGENLEDGER.SetFocus
        End If
        RS.Close
    End If
End Sub

Private Sub Combosubledger_GotFocus()
    If Trim(COMBOGENLEDGER.Text) = "" Then
        COMBOGENLEDGER.SetFocus
    End If
End Sub

Private Sub Combosubledger_LostFocus()
If Trim(Combosubledger.Text) <> "" Then
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        If RS.State = 1 Then
            RS.Close
        End If
        RS.Open "select * from sledger where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.Close
    Else
        Combosubledger.Text = ""
    End If
End If
End Sub

Private Sub AgCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim RS As New ADODB.Recordset
    RS.Open "select * from districts  where   " & stringyear & " and AGENTNAME= '" + AgCombo.Text + "'", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not RS.EOF Then
       Combosldistrictcode.Clear
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS(0)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    Combosldistrictcode.SetFocus
End If
End Sub

Private Sub Combosldistrictcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 
SendKeys "{TAB}"


End If


End Sub

Private Sub Commandreturn_Click()
''MainMenu.Toolbar1.Visible = True
    Unload Me
End Sub
Private Sub Commandshow_Click()
    Dim Ars  As New ADODB.Recordset
    Dim rs1  As New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim Balance As Double
    Set rs2 = New ADODB.Recordset
    Set rs3 = New ADODB.Recordset
    Set rs4 = New ADODB.Recordset
    CON.Execute "DELETE from rpttempindis1 where " & stringyear & ""
    If Trim(Combosldistrictcode.Text) <> "" Then
      If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
            MsgBox "invalid date"
            Exit Sub
      End If
      Balance = 0
     'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,createdby,createdon,fyear,setupid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C'as Vtype,CREDITA.SUBLEDGER as subleger , CREDITA.district AS DISTNAME, CREDITB.NETAMOUNT AS BNETAMT, CREDITB.BOOKCODE, BOOKS.GROUPCODE , " & UId & _
     " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) LEFT JOIN GROUPS ON BOOKS.GROUPCODE = GROUPS.groupname WHERE CREDITA.INVOICEDATE >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And CREDITA.INVOICEDATE<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.SUBLEDGER  In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "') AND CREDITA.district IN( select DISTCODE  from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE ='" + DWSalesReturn.Combosldistrictcode + "') ORDER BY CREDITA.INVOICEDATE"
     'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,createdby,createdon,fyear,setupid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'I'as Vtype,CREDITA.SUBLEDGER as subleger , CREDITA.district AS DISTNAME, CREDITB.NETAMOUNT AS BNETAMT, CREDITB.BOOKCODE, BOOKS.GROUPCODE," & UId & _
     " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from (CREDITA. LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) LEFT JOIN GROUPS ON BOOKS.GROUPCODE = GROUPS.groupname WHERE CREDITA.INVOICEDATE >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And CREDITA.INVOICEDATE<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.SUBLEDGER  In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "') AND  CREDITA.district IN( select DISTCODE  from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE ='" + DWSalesReturn.Combosldistrictcode + "') ORDER BY CREDITA.INVOICEDATE"
      
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,fyear,setupid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C'as Vtype,CREDITA.SUBLEDGER as subleger , CREDITA.district AS DISTNAME, CREDITB.NETAMOUNT AS BNETAMT, CREDITB.BOOKCODE, BOOKS.GROUPCODE," & UId & " as userid,'" & main.session & "'," & setupid & "" & _
      " FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE and CREDITB.fyear = BOOKS.fyear and CREDITB.setupid = BOOKS.setupid) ON CREDITA.INVOICENO = CREDITB.INVOICENO and CREDITB.fyear = credita.fyear and CREDITB.setupid = credita.setupid) LEFT JOIN GROUPS ON BOOKS.GROUPCODE = GROUPS.groupname WHERE  credita.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.SUBLEDGER  In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "') AND CREDITA.district IN( select DISTCODE  from Sledger WHERE  " & stringyear & " and " & _
      " gledger='SUNDRY DEBTORS' AND DISTCODE ='" + DWSalesReturn.Combosldistrictcode + "') ORDER BY CREDITA.INVOICEDATE"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,fyear,setupid)  SELECT CNF1B.CND AS VDATE, CNF1B.CNN AS VNO, 'N' AS Vtype, CNF1A.PSLD,SLEDGER.DISTCODE, CNF1B.A AS BNETAMT, " & main.UId & " AS userid,'" & main.session & "'," & setupid & "   FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN and cNF1B.setupid= cNF1A.setupid and cNF1B.fyear = cNF1A.fyear) ON SLEDGER.SUBLEDGER = CNF1A.PSLD and SLEDGER.setupid = CNF1A.setupid and SLEDGER.fyear = CNF1A.fyear  WHERE " & _
      " ((convert(smalldatetime,(CNF1B.CND),103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And  cnf1a.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and convert(smalldatetime,(CNF1B.CND),103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) ) AND ((CNF1A.PSLD) In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' and DISTCODE='" + DWSalesReturn.Combosldistrictcode + "')) AND ((CNF1B.GLD)='SALES RETURN'))  ORDER BY CNF1B.CND,CNF1B.CNN"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,fyear,setupid)  SELECT DNFB.DND AS VDATE, DNFB.DNN AS VNO, 'D' AS Vtype, DNFA.PSLD,SLEDGER.DISTCODE, DNFB.A AS BNETAMT,  " & main.UId & " AS userid,'" & main.session & "'," & setupid & "  FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN and DNFB.setupid= DNFA.setupid and DNFB.fyear = DNFA.fyear) ON SLEDGER.SUBLEDGER = DNFA.PSLD and SLEDGER.setupid = DNFA.setupid and SLEDGER.fyear = DNFA.fyear  WHERE ((convert(smalldatetime,(DNFB.DND),103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And  dnfa.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & "  and convert(smalldatetime,(DNFB.DND),103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) ) AND ((DNFA.PSLD) In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' and DISTCODE='" + DWSalesReturn.Combosldistrictcode + "')) AND ((DNFB.GLD)='SALES RETURN'))  ORDER BY DNFB.DND,DNFB.DNN"
 
      
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,createdby,createdon,fyear,setupid) SELECT CNF1A.CND AS VDATE, CNF1A.CNN AS VNO, 'N' as Vtype, CNF1A.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME,IIF(CNF1A.DC = 'D',NA,-NA) AS BNETAMT, " & UId & " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from CNF1A LEFT JOIN SLEDGER ON CNF1A.PSLD = SLEDGER.SUBLEDGER Where   CNF1A.PSLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "') AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "')   AND   convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CND,CNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,createdby,createdon,fyear,setupid) SELECT CNF1A.CND AS VDATE, CNF1A.CNN AS VNO, 'N' as Vtype, CNF1A.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, -NA AS BNETAMT, " & UId & " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from CNF1A LEFT JOIN SLEDGER ON CNF1A.PSLD = SLEDGER.SUBLEDGER Where   CNF1A.PSLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "') AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "')   AND   convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CND,CNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,createdby,createdon,fyear,setupid) SELECT CNF1B.CND AS VDATE, CNF1B.CNN AS VNO, 'N' as Vtype, CNF1B.SLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, -A AS BNETAMT, " & UId & " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from CNF1B LEFT JOIN SLEDGER ON CNF1B.SLD = SLEDGER.SUBLEDGER Where   CNF1B.SLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "') AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "')   AND   convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CND,CNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT,userid,createdby,createdon,fyear,setupid) SELECT DNFA.DND AS VDATE, DNFA.DNN AS VNO, 'D' as Vtype, DNFA.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, NA AS BNETAMT , " & UId & " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from DNFA LEFT JOIN SLEDGER ON DNFA.PSLD = SLEDGER.SUBLEDGER Where   DNFA.PSLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "') AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "')   AND   convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,createdby,createdon,fyear,setupid) SELECT DNFB.DND AS VDATE, DNFB.DNN AS VNO, 'D' as Vtype, DNFB.SLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, A AS BNETAMT, " & UId & " as userid  FROM DNFB LEFT JOIN SLEDGER ON DNFB.SLD = SLEDGER.SUBLEDGER Where DNFB.SLD In (select Subledger from Sledger WHERE Gledger='SUNDRY DEBTORS' AND DISTCODE ='" + DWSalesReturn.Combosldistrictcode + "') AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "')   AND  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN"
      
      main.reportname = "Dis. Sales"
      ViewlDisSalesRet.genreport
      PrintOption.Show
   Else
      If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
         MsgBox "invalid date"
         Exit Sub
      End If
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,createdby,createdon,fyear,setupid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C'as Vtype,CREDITA.SUBLEDGER as subleger , CREDITA.district AS DISTNAME, CREDITB.NETAMOUNT AS BNETAMT, CREDITB.BOOKCODE, BOOKS.GROUPCODE , " & UId & " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) LEFT JOIN GROUPS ON BOOKS.GROUPCODE = GROUPS.groupname WHERE CREDITA.INVOICEDATE >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And CREDITA.INVOICEDATE<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.SUBLEDGER  In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS  " & stringyear & " and WHERE AGENTNAME= '" & AgCombo.Text & "'" _
      '& " ))AND CREDITA.district IN( select DISTCODE  from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE  in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "')) ORDER BY CREDITA.INVOICEDATE"
     
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,fyear,setupid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C'as Vtype,CREDITA.SUBLEDGER as subleger , CREDITA.district AS DISTNAME, CREDITB.NETAMOUNT AS BNETAMT, CREDITB.BOOKCODE, BOOKS.GROUPCODE , " & UId & " as userid,'" & main.session & "'," & setupid & " from (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE and CREDITB.fyear = BOOKS.fyear and CREDITB.fyear = BOOKS.fyear and CREDITB.setupid = BOOKS.setupid)" & _
        " ON CREDITA.INVOICENO = CREDITB.INVOICENO and CREDITA.fyear = CREDITB.fyear and  CREDITA.setupid = CREDITB.setupid) WHERE  credita.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AGENTNAME= '" & AgCombo.Text & "' ORDER BY CREDITA.INVOICEDATE"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,createdby,createdon,fyear,setupid)  SELECT CNF1B.CND AS VDATE, CNF1B.CNN AS VNO, 'N' AS Vtype, CNF1A.SLD,SLEDGER.DISTCODE, CNF1B.A AS BNETAMT , 1 AS userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD  WHERE (((CNF1B.CND) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And (CNF1B.CND)<=convert(smalldatetime,'" & trim(date2.text) & "',103) ) AND ((CNF1A.PSLD) In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' and AGENTNAME= '" & AgCombo.Text & "')) AND ((CNF1B.GLD)='SALES RETURN'))  ORDER BY CNF1B.CND,CNF1B.CNN"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,fyear,setupid)  SELECT CNF1B.CND AS VDATE, CNF1B.CNN AS VNO, 'N' AS Vtype, CNF1A.PSLD,SLEDGER.DISTCODE, CNF1B.A AS BNETAMT , " & UId & " AS userid,'" & main.session & "'," & setupid & " from SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN and CNF1B.fyear = CNF1A.fyear and CNF1B.setupid = CNF1A.setupid) ON SLEDGER.SUBLEDGER = CNF1A.PSLD  WHERE ((convert(smalldatetime,CNF1B.CND,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And  cnf1b.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and convert(smalldatetime,CNF1B.CND,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) ) AND ((CNF1A.PSLD) In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' And DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & "))) AND ((CNF1B.GLD)='SALES RETURN'))  ORDER BY CNF1B.CND,CNF1B.CNN"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,fyear,setupid)  SELECT DNFB.DND AS VDATE, DNFB.DNN AS VNO, 'D' AS Vtype, DNFA.PSLD,SLEDGER.DISTCODE, DNFB.A AS BNETAMT,  " & UId & " AS userid,'" & main.session & "'," & setupid & "  FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN and DNFB.fyear = DNFA.fyear and DNFB.setupid = DNFA.setupid) ON SLEDGER.SUBLEDGER = DNFA.PSLD  WHERE ((convert(smalldatetime,DNFB.DND,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And  dnfa.fyear='" & main.session & "' and dnfa.setupid='" & main.setupid & "' and convert(smalldatetime,DNFB.DND,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) ) AND ((DNFA.PSLD) In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' And DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & "))) AND ((DNFB.GLD)='SALES RETURN'))  ORDER BY DNFB.DND,DNFB.DNN"
      
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT,userid,createdby,createdon,fyear,setupid) SELECT CNF1A.CND AS VDATE, CNF1A.CNN AS VNO, 'N' as Vtype, CNF1A.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, -NA AS BNETAMT , " & UId & " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from CNF1A LEFT JOIN SLEDGER ON CNF1A.PSLD = SLEDGER.SUBLEDGER Where   CNF1A.PSLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "') ) AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'))  AND " & _
      '"  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CND,CNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT,userid,createdby,createdon,fyear,setupid) SELECT CNF1B.CND AS VDATE, CNF1B.CNN AS VNO, 'N' as Vtype, CNF1B.SLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, -A AS BNETAMT , " & UId & " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from CNF1B LEFT JOIN SLEDGER ON CNF1B.SLD = SLEDGER.SUBLEDGER Where   CNF1B.SLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "') ) AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'))  AND " & _
      '"  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CND,CNN"
      
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT,userid,createdby,createdon,fyear,setupid) SELECT DNFA.DND AS VDATE, DNFA.DNN AS VNO, 'D' as Vtype, DNFA.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, NA AS BNETAMT , " & UId & " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from DNFA LEFT JOIN SLEDGER ON DNFA.PSLD = SLEDGER.SUBLEDGER Where   DNFA.PSLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "') ) AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'))  AND " & _
      '"  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT,userid,createdby,createdon,fyear,setupid) SELECT DNFB.DND AS VDATE, DNFB.DNN AS VNO, 'D' as Vtype, DNFB.SLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, A AS BNETAMT , " & UId & " as userid,'" &  main.username &  "','" & now & "','" &  main.session & "'," & setupid & " from DNFB LEFT JOIN SLEDGER ON DNFB.SLD = SLEDGER.SUBLEDGER Where   DNFB.SLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "') ) AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'))  AND " & _
      '"  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN"
      Balance = 0
      main.reportname = "Dis. Sales"
      ViewlDisSalesRet.genreport
      PrintOption.Show
  End If






'''''
'''''
'''''
'''''If Trim(Combosldistrictcode.Text) <> "" Then
'''''      If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
'''''            MsgBox "invalid date"
'''''            Exit Sub
'''''      End If
'''''      BALANCE = 0
'''''
''''''      If rs2.State = 1 Then rs2.Close
''''''
''''''      rs1.Open "select * from Sledger WHERE " & stringyear & " and gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWSalesReturn.Combosldistrictcode + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
''''''      If rs1.RecordCount > 0 Then
''''''         Do While Not rs1.EOF
''''''             If rs2.State = 1 Then rs2.Close
''''''             rs2.Open "SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO, SLEDGER.SUBLEDGER AS SUBLEGER, DISTRICTS.DISTRICTNAME AS DISTNAME, CREDITB.NETAMOUNT AS BNETAMT, CREDITB.BOOKCODE, BOOKS.GROUPCODE FROM (SLEDGER LEFT JOIN DISTRICTS ON SLEDGER.DISTCODE = DISTRICTS.DISTRICTNAME) RIGHT JOIN (GROUPS RIGHT JOIN ((CREDITB LEFT JOIN CREDITA ON CREDITB.INVOICENO = CREDITA.INVOICENO) INNER JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON GROUPS.groupcode = BOOKS.GROUPCODE) ON SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER Where " & stringyear & " and SLEDGER.SUBLEDGER = '" & rs1!SUBLEDGER & "' AND  DISTRICTS.DISTRICTNAME = '" & rs1!Distcode & "' AND   CREDITA.INVOICEDATE  >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and CREDITA.INVOICEDATE <=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CREDITA.INVOICEDATE, CREDITA.INVOICENO", CON, adOpenStatic, adCmdText
''''''
''''''             'FOR CREDITB NOTE
''''''             If rs3.State = 1 Then rs3.Close
''''''            'rs3.Open "SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO, SLEDGER.SUBLEDGER AS SUBLEGER, DISTRICTS.DISTRICTNAME AS DISTNAME, CREDITB.NETAMOUNT AS BNETAMT, CREDITB.BOOKCODE, BOOKS.GROUPCODE FROM GROUPS RIGHT JOIN ((((SLEDGER LEFT JOIN DISTRICTS ON SLEDGER.DISTCODE = DISTRICTS.DISTRICTNAME) RIGHT JOIN CREDITA ON SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) LEFT JOIN CREDITB ON CREDITA.INVOICENO = CREDITB.INVOICENO) LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON GROUPS.groupcode = BOOKS.GROUPCODE  Where " & stringyear & " and SLEDGER.SUBLEDGER = '" & rs1!SUBLEDGER & "' AND  DISTRICTS.DISTRICTNAME = '" & rs1!DISTCODE & "'ORDER BY CREDITA.INVOICEDATE, CREDITA.INVOICENO", con, adOpenStatic, adLockOptimistic, adCmdText
''''''             rs3.Open "SELECT DNFA.DND AS VDATE, DNFA.DNN AS VNO, DNFA.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, NA AS BNETAMT  FROM DNFA LEFT JOIN SLEDGER ON DNFA.PSLD = SLEDGER.SUBLEDGER Where   " & stringyear & " and DNFA.PSLD = '" & rs1!SUBLEDGER & "' AND SLEDGER.DISTCODE= '" & rs1!Distcode & "'  AND   convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN", CON, adOpenStatic, adLockOptimistic, adCmdText
''''''             If rs2.RecordCount > 0 Then
''''''                rs2.MoveFirst
''''''                Do While Not rs2.EOF
''''''                   CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE) values( '" & rs2!vdate & "', " & rs2!vno & ",'C','" & rs2!SUBLEGER & "','" & rs2!DISTNAME & "'," & rs2!BNETAMT & ",'" & rs2!Bookcode & "','" & rs2!groupcode & "')"
''''''                   rs2.MoveNext
''''''                Loop
''''''            End If
''''''            If rs3.RecordCount > 0 Then
''''''                rs3.MoveFirst
''''''                Do While Not rs3.EOF
''''''                     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ) values( '" & rs3!vdate & "', " & rs3!vno & ",'D','" & rs3!SUBLEGER & "','" & rs3!DISTNAME & "'," & rs3!BNETAMT & ")"
''''''                     rs3.MoveNext
''''''                Loop
''''''            End If
''''''            rs1.MoveNext
''''''       Loop
'''''   End If
'''''
'''''   If rs1.State = 1 Then rs1.Close
'''''   rs1.Open "treport", CON, adOpenKeyset, adLockReadOnly, adcmdtext
'''''   rs1.Close
'''''   If rs.State = 1 Then
'''''        rs.Close
'''''   End If
'''''    main.reportname = "Dis. Sales"
'''''    ViewlDisSalesRet.genreport
'''''    PrintOption.Show
'''''    PrintOption.Show
'''''    ViewlDisSalesRet.Top = 0
'''''    ViewlDisSalesRet.Left = 0
'''''    ViewlDisSalesRet.Show
'''''Else
'''''
'''''    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
'''''        MsgBox "invalid date"
'''''        Exit Sub
'''''    End If
'''''    BALANCE = 0
'''''    If Ars.State = 1 Then Ars.Close
'''''    Ars.Open "SELECT DISTINCT * FROM districts WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'", CON, adOpenStatic, adLockReadOnly, adCmdText
'''''    If Ars.RecordCount > 0 Then
'''''          Do While Not Ars.EOF
'''''             If rs1.State = 1 Then rs1.Close
'''''             rs1.Open "select * from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + Ars!DISTRICTNAME + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
'''''             If rs1.RecordCount > 0 Then
'''''             Do While Not rs1.EOF
'''''
'''''                If rs2.State = 1 Then rs2.Close
'''''                rs2.Open "SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO, SLEDGER.SUBLEDGER AS SUBLEGER, DISTRICTS.DISTRICTNAME AS DISTNAME, CREDITB.NETAMOUNT AS BNETAMT, CREDITB.BOOKCODE, BOOKS.GROUPCODE FROM (SLEDGER LEFT JOIN DISTRICTS ON SLEDGER.DISTCODE = DISTRICTS.DISTRICTNAME) RIGHT JOIN (GROUPS RIGHT JOIN ((CREDITB LEFT JOIN CREDITA ON CREDITB.INVOICENO = CREDITA.INVOICENO) INNER JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON GROUPS.groupcode = BOOKS.GROUPCODE) ON SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER Where SLEDGER.SUBLEDGER = '" & rs1!SUBLEDGER & "' AND  DISTRICTS.DISTRICTNAME = '" & rs1!DISTCODE & "'  AND   CREDITA.INVOICEDATE  >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and CREDITA.INVOICEDATE <=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY CREDITA.INVOICEDATE, CREDITA.INVOICENO", CON, adOpenStatic, adLockOptimistic, adCmdText
'''''                 FOR CREDIT NOTE ITEM
'''''                If rs3.State = 1 Then rs3.Close
'''''                rs3.Open "SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO, SLEDGER.SUBLEDGER AS SUBLEGER, DISTRICTS.DISTRICTNAME AS DISTNAME, CREDITB.NETAMOUNT AS BNETAMT, CREDITB.BOOKCODE, BOOKS.GROUPCODE FROM GROUPS RIGHT JOIN ((((SLEDGER LEFT JOIN DISTRICTS ON SLEDGER.DISTCODE = DISTRICTS.DISTRICTNAME) RIGHT JOIN CREDITA ON SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) LEFT JOIN CREDITB ON CREDITA.INVOICENO = CREDITB.INVOICENO) LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON GROUPS.groupcode = BOOKS.GROUPCODE  Where SLEDGER.SUBLEDGER = '" & rs1!SUBLEDGER & "' AND  DISTRICTS.DISTRICTNAME = '" & rs1!DISTCODE & "'ORDER BY CREDITA.INVOICEDATE, CREDITA.INVOICENO", CON, adOpenStatic, adLockOptimistic, adCmdText
'''''                rs3.Open "SELECT DNFA.DND AS VDATE, DNFA.DNN AS VNO, DNFA.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME,NA AS BNETAMT  FROM DNFA LEFT JOIN SLEDGER ON DNFA.PSLD = SLEDGER.SUBLEDGER Where   DNFA.PSLD = '" & rs1!SUBLEDGER & "' AND SLEDGER.DISTCODE= '" & rs1!DISTCODE & "' AND   convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN", CON, adOpenStatic, adLockOptimistic, adCmdText
'''''
'''''                If rs2.RecordCount > 0 Then
'''''                    rs2.MoveFirst
'''''                    Do While Not rs2.EOF
'''''                         CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE) values( '" & rs2!vdate & "', " & rs2!vno & ",'C','" & rs2!SUBLEGER & "','" & rs2!DISTNAME & "'," & rs2!BNETAMT & ",'" & rs2!Bookcode & "','" & rs2!groupcode & "')"
'''''                         rs2.MoveNext
'''''                    Loop
'''''                End If
'''''
'''''                If rs3.RecordCount > 0 Then
'''''                    rs3.MoveFirst
'''''                    Do While Not rs3.EOF
'''''                                  CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ) values( '" & rs3!vdate & "', " & rs3!vno & ",'D','" & rs3!SUBLEGER & "','" & rs3!DISTNAME & "'," & rs3!BNETAMT & ")"
'''''                                  rs3.MoveNext
'''''                    Loop
'''''                End If
'''''                rs1.MoveNext
'''''          Loop
'''''     End If
'''''     Ars.MoveNext
'''''     Loop
''''' End If
'''''    If rs1.State = 1 Then rs1.Close
'''''    rs1.Open "treport", CON, adOpenKeyset, adLockReadOnly, adcmdtext
'''''    rs1.Close
'''''    If rs.State = 1 Then
'''''        rs.Close
'''''    End If
'''''
'''''    main.reportname = "Dis. Sales"
'''''    ViewlDisSalesRet.genreport
'''''    PrintOption.Show
'''''
'''''    ViewlDisSalesRet.Top = 0
'''''    ViewlDisSalesRet.Left = 0
'''''    ViewlDisSalesRet.Show
'''''End If
End Sub
Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}"
End If
End Sub

Private Sub date1_LostFocus()
    If Trim(date1.Text) <> "" Then
        If Not checkdate(Trim(date1.Text), date1) Then
            date1.SetFocus
        End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}"
End If
End Sub

Private Sub date2_LostFocus()
    If Trim(date2.Text) <> "" Then
        If Not checkdate(Trim(date2.Text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop
Me.TOP = 0
Me.Left = 0
'Set CON = New ADODB.Connection
Set RS = New ADODB.Recordset
'    With CON
   '     .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
  '      .Open
  '  End With
    CNSetup
    RS.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
    RS.Close
    RS.Open "select * from DISTRICTS where  " & stringyear & "  order by DISTRICTNAME", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS!DISTRICTNAME
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    RS.Open "select Agentname  from AgentMaster where " & stringyear & " order by AgentNAME", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
        If IsNull(RS!agentname) = False Then
            Me.AgCombo.AddItem RS!agentname
        End If
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close




End Sub

