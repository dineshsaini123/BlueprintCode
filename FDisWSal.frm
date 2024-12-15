VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DWsales 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4755
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox AgCombo 
      Height          =   315
      Left            =   2550
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   690
      Width           =   3885
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   4380
      TabIndex        =   5
      Top             =   3030
      Width           =   1455
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   2640
      TabIndex        =   4
      Top             =   3060
      Width           =   1545
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   2550
      TabIndex        =   2
      Top             =   1650
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
      Left            =   2550
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1170
      Width           =   3885
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   4440
      TabIndex        =   3
      Top             =   1650
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
      Left            =   1230
      TabIndex        =   9
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   315
      Left            =   3930
      TabIndex        =   8
      Top             =   1710
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From The Date"
      Height          =   195
      Left            =   1110
      TabIndex        =   7
      Top             =   1740
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Agent Name"
      Height          =   195
      Left            =   1290
      TabIndex        =   6
      Top             =   690
      Width           =   885
   End
End
Attribute VB_Name = "DWsales"
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
        RS.Open "select * from gledger where " & stringyear & " and slf=0", CON, adOpenKeyset, adLockReadOnly, adCmdText
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
    RS.Open "select  * from DISTRICTS  where   " & stringyear & " and AGENTNAME= '" + AgCombo.Text + "'", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
        Combosldistrictcode.Clear
    If Not RS.EOF Then
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

Private Sub AgCombo_LostFocus()
  Dim RS As New ADODB.Recordset

 If AgCombo.Text <> "" Then
    RS.Open "select  * from DISTRICTS  where   " & stringyear & " and AGENTNAME= '" + AgCombo.Text + "'", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not RS.EOF Then
       Combosldistrictcode.Clear
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS(1)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    Combosldistrictcode.SetFocus
  Else
  
     
     RS.Open "select  * from DISTRICTS where  " & stringyear & "", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
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
  CON.Execute "DELETE from rpttempindis1 where  " & stringyear & " and userid=" & main.UId
  If Trim(Combosldistrictcode.Text) <> "" Then
      If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
            MsgBox "invalid date"
            Exit Sub
      End If
      Balance = 0
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,fyear,setupid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I'as Vtype,INVOICEA.SUBLEDGER as subleger , INVOICEA.district AS DISTNAME, INVOICEB.NETAMOUNT AS BNETAMT, INVOICEB.BOOKCODE, BOOKS.GROUPCODE," & UId & " as userid,'" & main.session & "'," & main.setupid & " FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE and INVOICEB.fyear = BOOKS.fyear) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO and INVOICEA.fyear = INVOICEB.fyear and INVOICEA.setupid = INVOICEB.setupid) LEFT JOIN GROUPS ON BOOKS.GROUPCODE = GROUPS.groupname WHERE  " & stringyear & " and " &
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,fyear,setupid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I'as Vtype,INVOICEA.SUBLEDGER as subleger , INVOICEA.district AS DISTNAME, INVOICEB.NETAMOUNT AS BNETAMT, INVOICEB.BOOKCODE, BOOKS.GROUPCODE," & UId & " as userid,'" & main.session & "'," & main.setupid & " FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE and INVOICEB.fyear = BOOKS.fyear and INVOICEB.setupid = BOOKS.setupid) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO and INVOICEA.fyear = INVOICEB.fyear and INVOICEA.setupid = INVOICEB.setupid) WHERE  invoicea.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and " & _
      "convert(smalldatetime,INVOICEA.INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.SUBLEDGER  In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "') AND INVOICEA.district IN( select DISTCODE  from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' AND DISTCODE ='" + DWsales.Combosldistrictcode + "') ORDER BY INVOICEA.INVOICEDATE"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,fyear,setupid)  SELECT CNF1B.CND AS VDATE, CNF1B.CNN AS VNO, 'C' AS Vtype, CNF1A.PSLD,SLEDGER.DISTCODE, CNF1B.A AS BNETAMT, " & UId & " as userid ,'" & main.session & "'," & main.setupid & "  FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN and CNF1B.fyear = CNF1A.fyear and CNF1B.setupid = CNF1A.setupid) ON SLEDGER.SUBLEDGER = CNF1A.PSLD and SLEDGER.fyear = CNF1A.fyear  WHERE ((cnf1a.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & ") and (convert(smalldatetime,CNF1B.CND,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,CNF1B.CND,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103)) AND ((CNF1A.PSLD) In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' and DISTCODE='" + DWsales.Combosldistrictcode + "')) AND ((CNF1B.GLD)='SALES'))  ORDER BY CNF1B.CND,CNF1B.CNN"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,fyear,setupid)  SELECT DNFB.DND AS VDATE, DNFB.DNN AS VNO, 'D' AS Vtype, DNFA.PSLD,SLEDGER.DISTCODE, DNFB.A AS BNETAMT,  " & UId & " as userid ,'" & main.session & "'," & main.setupid & " FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN and DNFB.fyear = DNFA.fyear and DNFB.setupid = DNFA.setupid) ON SLEDGER.SUBLEDGER = DNFA.PSLD and SLEDGER.fyear = DNFA.fyear WHERE ((dnfa.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & ") and (convert(smalldatetime,DNFB.DND,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,DNFB.DND,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103)) AND (DNFA.PSLD In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' and DISTCODE='" + DWsales.Combosldistrictcode + "')) AND ((DNFB.GLD)='SALES'))  ORDER BY DNFB.DND,DNFB.DNN"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,fyear,setupid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S'as Vtype,CASHA.SUBLEDGER as subleger , CASHA.district AS DISTNAME, CASHB.NETAMOUNT AS BNETAMT, CASHB.BOOKCODE, BOOKS.GROUPCODE, " & UId & " as userid,'" & main.session & "'," & main.setupid & "" & _
      " FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE and CASHB.fyear = BOOKS.fyear ) ON CASHA.INVOICENO = CASHB.INVOICENO and CASHA.fyear = CASHB.fyear and CASHA.setupid = CASHB.setupid) WHERE casha.fyear='" & main.session & "' and casha.setupid=" & main.setupid & " and convert(smalldatetime,CASHA.INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.SUBLEDGER  In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS') AND CASHA.district  In (select DISTCODE from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "') ORDER BY CASHA.INVOICEDATE"
'     " FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE and CASHB.fyear = BOOKS.fyear ) ON CASHA.INVOICENO = CASHB.INVOICENO and CASHA.fyear = CASHB.fyear and CASHA.setupid = CASHB.setupid) LEFT JOIN GROUPS ON BOOKS.GROUPCODE = GROUPS.groupname WHERE  " & stringyear & " and CASHA.INVOICEDATE >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And CASHA.convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.SUBLEDGER  In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS') AND CASHA.district  In (select DISTCODE from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "') ORDER BY CASHA.INVOICEDATE"
      
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,createdby,createdon,fyear,setupid) SELECT CNF1A.CND AS VDATE, CNF1A.CNN AS VNO, 'C' as Vtype, CNF1A.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, NA AS BNETAMT, " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & "  FROM CNF1A LEFT JOIN SLEDGER ON CNF1A.PSLD = SLEDGER.SUBLEDGER Where   CNF1A.PSLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "') AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "')   AND   convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CND,CNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,createdby,createdon,fyear,setupid) SELECT CNF1B.CND AS VDATE, CNF1B.CNN AS VNO, 'C' as Vtype, CNF1B.SLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, A AS BNETAMT, " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & "  FROM CNF1B LEFT JOIN SLEDGER ON CNF1B.SLD = SLEDGER.SUBLEDGER Where   CNF1B.SLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "') AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "')   AND   convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CND,CNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,createdby,createdon,fyear,setupid) SELECT DNFA.DND AS VDATE, DNFA.DNN AS VNO, 'D' as Vtype, DNFA.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, -NA AS BNETAMT, " & UId & " as userid ,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & " FROM DNFA LEFT JOIN SLEDGER ON DNFA.PSLD = SLEDGER.SUBLEDGER Where   DNFA.PSLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "') AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "')   AND   convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,createdby,createdon,fyear,setupid) SELECT DNFB.DND AS VDATE, DNFB.DNN AS VNO, 'D' as Vtype, DNFB.SLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, -A AS BNETAMT, " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & "  FROM DNFB LEFT JOIN SLEDGER ON DNFB.SLD = SLEDGER.SUBLEDGER Where DNFB.SLD In (select Subledger from Sledger WHERE Gledger='SUNDRY DEBTORS' AND DISTCODE ='" + DWsales.Combosldistrictcode + "') AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "')   AND  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,createdby,createdon,fyear,setupid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S'as Vtype,CASHA.SUBLEDGER as subleger , CASHA.district AS DISTNAME, CASHB.NETAMOUNT AS BNETAMT, CASHB.BOOKCODE, BOOKS.GROUPCODE, " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & " FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) LEFT JOIN GROUPS ON BOOKS.GROUPCODE = GROUPS.groupname WHERE CASHA.INVOICEDATE >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And " & _
      '"CASHA.convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.SUBLEDGER  In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS') AND CASHA.district  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE='" + DWsales.Combosldistrictcode + "') ORDER BY CASHA.INVOICEDATE"
      main.reportname = "Dis. Sales"
      viewlDisSales.genreport
      PrintOption.Show
  Else
      If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
         MsgBox "invalid date"
         Exit Sub
      End If
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,createdby,createdon,fyear,setupid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I'as Vtype,INVOICEA.SUBLEDGER as subleger , INVOICEA.district AS DISTNAME, INVOICEB.NETAMOUNT AS BNETAMT, INVOICEB.BOOKCODE, BOOKS.GROUPCODE, " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & " FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) LEFT JOIN GROUPS ON BOOKS.GROUPCODE = GROUPS.groupname WHERE INVOICEA.INVOICEDATE >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And " & _
      ' INVOICEA.INVOICEDATE<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.SUBLEDGER  In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'" _
      & " ))AND INVOICEA.district IN( select DISTCODE  from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE  in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "')) ORDER BY INVOICEA.INVOICEDATE"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,fyear,setupid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I'as Vtype,INVOICEA.SUBLEDGER as subleger , INVOICEA.district AS DISTNAME, INVOICEB.NETAMOUNT AS BNETAMT, INVOICEB.BOOKCODE, BOOKS.GROUPCODE, " & UId & " as userid,'" & main.session & "'," & main.setupid & " FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE and INVOICEB.fyear = BOOKS.fyear and INVOICEB.setupid = BOOKS.setupid)" & _
      " ON INVOICEA.INVOICENO = INVOICEB.INVOICENO and INVOICEA.fyear = INVOICEB.fyear and INVOICEA.setupid = INVOICEB.setupid) WHERE  invoicea.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and convert(smalldatetime,INVOICEA.INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.Text & "'" & " ORDER BY INVOICEA.INVOICEDATE"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,fyear,setupid)  SELECT CNF1B.CND AS VDATE, CNF1B.CNN AS VNO, 'C' AS Vtype, CNF1A.PSLD,SLEDGER.DISTCODE, CNF1B.A AS BNETAMT , " & UId & " as userid,'" & main.session & "'," & main.setupid & " FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN and CNF1B.fyear = CNF1A.fyear and  CNF1B.setupid= CNF1A.setupid) ON SLEDGER.SUBLEDGER = CNF1A.PSLD and SLEDGER.fyear = CNF1A.fyear and SLEDGER.setupid = CNF1A.setupid  WHERE ((CNF1A.fyear='" & main.session & "' and CNF1A.setupid=" & main.setupid & " and " & _
      " convert(smalldatetime,CNF1B.CND,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CNF1B.CND,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103)) AND ((CNF1A.PSLD) In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS' And DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & "))) AND ((CNF1B.GLD)='SALES'))  ORDER BY CNF1B.CND,CNF1B.CNN"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT ,userid,fyear,setupid)  SELECT DNFB.DND AS VDATE, DNFB.DNN AS VNO, 'D' AS Vtype, DNFA.PSLD,SLEDGER.DISTCODE, DNFB.A AS BNETAMT,  " & UId & " as userid,'" & main.session & "'," & main.setupid & "  FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN and dNFB.fyear = dNFA.fyear and  dNFB.setupid= dNFA.setupid) ON SLEDGER.SUBLEDGER = DNFA.PSLD and SLEDGER.fyear = dNFA.fyear and SLEDGER.setupid = dNFA.setupid " & _
      "WHERE ((dNFA.fyear='" & main.session & "' and dNFA.setupid=" & main.setupid & " and convert(smalldatetime,DNFB.DND,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,DNFB.DND,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) ) AND ((DNFA.PSLD) In (select Subledger from Sledger WHERE  " & stringyear & " and gledger='SUNDRY DEBTORS'And DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & "))) AND ((DNFB.GLD)='SALES'))  ORDER BY DNFB.DND,DNFB.DNN"
      CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,fyear,setupid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S'as Vtype,CASHA.SUBLEDGER as subleger , CASHA.district AS DISTNAME, CASHB.NETAMOUNT AS BNETAMT, CASHB.BOOKCODE, BOOKS.GROUPCODE , " & UId & " as userid,'" & main.session & "'," & main.setupid & " FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE and CASHB.fyear = BOOKS.fyear and CASHB.setupid = BOOKS.setupid) ON CASHA.INVOICENO = CASHB.INVOICENO and cashA.fyear = cashB.fyear and cashA.setupid = cashB.setupid) WHERE casha.fyear='" & main.session & "' and casha.setupid=" & main.setupid & " and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.AGENTNAME= '" & AgCombo.Text & "' ORDER BY CASHA.INVOICEDATE"
     
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT,userid,createdby,createdon,fyear,setupid) SELECT CNF1A.CND AS VDATE, CNF1A.CNN AS VNO, 'C' as Vtype, CNF1A.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, NA AS BNETAMT , " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & " FROM CNF1A LEFT JOIN SLEDGER ON CNF1A.PSLD = SLEDGER.SUBLEDGER Where   CNF1A.PSLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "') ) AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'))  AND  " & _
      '"convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CND,CNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT,userid,createdby,createdon,fyear,setupid) SELECT CNF1B.CND AS VDATE, CNF1B.CNN AS VNO, 'C' as Vtype, CNF1B.SLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, A AS BNETAMT , " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & " FROM CNF1B LEFT JOIN SLEDGER ON CNF1B.SLD = SLEDGER.SUBLEDGER Where   CNF1B.SLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "') ) AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'))  AND " & _
      '"  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CND,CNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT,userid,createdby,createdon,fyear,setupid) SELECT DNFA.DND AS VDATE, DNFA.DNN AS VNO, 'D' as Vtype, DNFA.PSLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, -NA AS BNETAMT , " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & " FROM DNFA LEFT JOIN SLEDGER ON DNFA.PSLD = SLEDGER.SUBLEDGER Where   DNFA.PSLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "') ) AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'))  AND " & _
      '"  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN"
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT,userid,createdby,createdon,fyear,setupid) SELECT DNFB.DND AS VDATE, DNFB.DNN AS VNO, 'D' as Vtype, DNFB.SLD AS SUBLEGER, SLEDGER.DISTCODE AS DISTNAME, -A AS BNETAMT , " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & " FROM DNFB LEFT JOIN SLEDGER ON DNFB.SLD = SLEDGER.SUBLEDGER Where   DNFB.SLD In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "') ) AND SLEDGER.DISTCODE  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in (SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "'))  AND " & _
      '"  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY DND,DNN"
     
      'CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, DISTNAME, BNETAMT, BOOKCODE, GROUPCODE,userid,createdby,createdon,fyear,setupid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S'as Vtype,CASHA.SUBLEDGER as subleger , CASHA.district AS DISTNAME, CASHB.NETAMOUNT AS BNETAMT, CASHB.BOOKCODE, BOOKS.GROUPCODE , " & UId & " as userid,'" & main.username & "','" & Now & ",'" & main.session & "'," & main.setupid & " FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) LEFT JOIN GROUPS ON BOOKS.GROUPCODE = GROUPS.groupname WHERE CASHA.INVOICEDATE >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And CASHA.INVOICEDATE<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.SUBLEDGER  In (select Subledger from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in ( SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & _
       AgCombo.Text & "')) AND CASHA.district  In (select DISTCODE from Sledger WHERE gledger='SUNDRY DEBTORS' AND DISTCODE in ( SELECT  DISTRICTNAME  FROM DISTRICTS WHERE  " & stringyear & " and AGENTNAME= '" & AgCombo.Text & "') ) ORDER BY CASHA.INVOICEDATE"
      
      
      Balance = 0
      main.reportname = "Dis. Sales"
      viewlDisSales.genreport
      PrintOption.Show
  End If
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
 '       .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
 '       .Open
 '   End With
    CNSetup
    RS.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
    RS.Close
    RS.Open "select * from DISTRICTS where  " & stringyear & " order by DISTRICTNAME", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS!DISTRICTNAME
          
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.Close
    RS.Open "select  Agentname  from AgentMaster where " & stringyear & " order by AgentNAME", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
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

