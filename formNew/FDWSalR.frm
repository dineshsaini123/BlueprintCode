VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form DWSalesReturn 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox AgCombo 
      Height          =   315
      Left            =   1860
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3885
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   3480
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   1860
      TabIndex        =   4
      Top             =   2640
      Width           =   1545
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   1860
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
      Left            =   1860
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3885
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   3750
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
   Begin Crystal.CrystalReport cr1 
      Left            =   900
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Dis. Sales"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "District Name"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   1140
      Width           =   1065
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   315
      Left            =   3240
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
      Left            =   480
      TabIndex        =   6
      Top             =   630
      Width           =   1065
   End
End
Attribute VB_Name = "DWSalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As Recordset
Private Sub COMBOGENLEDGER_LostFocus()
COMBOGENLEDGER = UCase(COMBOGENLEDGER)
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        RS.Open "select * from gledger where " & stringyear & " and slf=FALSE", CON, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.BOF Then
                COMBOGENLEDGER.SetFocus
        End If
        RS.close
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
            RS.close
        End If
        RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", CON, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.close
    Else
        Combosubledger.Text = ""
    End If
End If
End Sub

Private Sub AgCombo_Click()
    If AgCombo.Text = "" Then Exit Sub
    Dim RS As New ADODB.Recordset
    RS.Open "select * from districts  where  " & stringyear & " and AGENTNAME= '" + AgCombo.Text + "'", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not RS.EOF Then
       Combosldistrictcode.Clear
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS(0)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    'Combosldistrictcode.SetFocus

End Sub

Private Sub AgCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim RS As New ADODB.Recordset
    RS.Open "select * from districts  where  " & stringyear & " and AGENTNAME= '" + AgCombo.Text + "'", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not RS.EOF Then
       Combosldistrictcode.Clear
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS(0)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    Combosldistrictcode.SetFocus
End If
End Sub

Private Sub Combosldistrictcode_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   SendKeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
      SendKeys "{Down}"
      SendKeys "{tab}"
End If



End Sub

Private Sub Commandreturn_Click()
    Unload Me
End Sub
Private Sub Commandshow_Click()



Commandshow.Enabled = False
  Dim Ars  As New ADODB.Recordset
  Dim rs1  As New ADODB.Recordset
  Dim rs2 As ADODB.Recordset
  Dim trs9 As New ADODB.Recordset
  Dim rs3 As ADODB.Recordset
  Dim rs4 As ADODB.Recordset
  Dim rs5 As New ADODB.Recordset
  Dim Balance As Double
  Set rs2 = New ADODB.Recordset
  Set rs3 = New ADODB.Recordset
  Set rs4 = New ADODB.Recordset
  Dim trs As New ADODB.Recordset
  CON.Execute "Delete  from rpttempindis1"
  If trs9.State = 1 Then trs9.close
  trs9.Open "Select *  from groups where " & stringyear & " and group1=1 order BY groupcode", CON, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g1 = "( " & trs9!groupcode & " )"
    End If
    If trs9.State = 1 Then trs9.close
    trs9.Open "Select *  from groups where " & stringyear & " and group2=1 order BY groupcode", CON, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g2 = "( " & trs9!groupcode & " )"
    End If
    
    If trs9.State = 1 Then trs9.close
    trs9.Open "Select *  from groups where " & stringyear & " and group3=1 order BY groupcode", CON, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g3 = "( " & trs9!groupcode & " )"
    End If
    
    If trs9.State = 1 Then trs9.close
    trs9.Open "Select *  from groups where " & stringyear & " and group4=1 order BY groupcode", CON, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g4 = "( " & trs9!groupcode & " )"
    End If

    
    
    If trs.State = 1 Then trs.close
    CON.Execute "Delete from treport where " & stringyear
    trs.Open "select * from treport where " & stringyear, CON, adOpenDynamic, adLockOptimistic, adCmdTable
    xstr = DWSalesReturn.date1 + "  To  " + DWSalesReturn.date2
    trs.AddNew
    trs!Text = xstr
    trs!Genledger = g1
    trs!SUBLEDGER = g2
    trs!narration = g3
    trs!header = g4
    
    trs.update
    
  If AgCombo.Text = "" And Trim(Combosldistrictcode.Text) <> "" Then
     If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
            MsgBox "invalid date"
            Exit Sub
     End If
     Balance = 0
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'I',CREDITA.SUBLEDGER as subleger ,CREDITA.agentname, CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) " & _
     "where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AgentName= '" & AgCombo.Text & "' and  CREDITA.district='" + DWSalesReturn.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER, CREDITA.AGENTNAME,CREDITA.district, BOOKS.GROUPCODE"
     
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'I',CREDITA.SUBLEDGER as subleger ,CREDITA.agentname, CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) " & _
     "where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND   CREDITA.district='" + DWSalesReturn.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER, CREDITA.AGENTNAME,CREDITA.district, BOOKS.GROUPCODE"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'I',CREDITA.SUBLEDGER as subleger , CREDITA.agentname,CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) " & _
     "where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND   CREDITA.district='" + DWSalesReturn.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER, CREDITA.AGENTNAME,CREDITA.district, BOOKS.GROUPCODE"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'I',CREDITA.SUBLEDGER as subleger , CREDITA.agentname,CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) " & _
     "where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND   CREDITA.district='" + DWSalesReturn.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER, CREDITA.AGENTNAME,CREDITA.district, BOOKS.GROUPCODE"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'I',CREDITA.SUBLEDGER as subleger , CREDITA.agentname,CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) " & _
     "where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND   CREDITA.district='" + DWSalesReturn.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER, CREDITA.AGENTNAME,CREDITA.district, BOOKS.GROUPCODE"
    
    
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,-sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND CNF1B.GLD='SALES RETURN' and  SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND CNF1B.GLD='SALES RETURN' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND CNF1B.GLD='SALES RETURN' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND CNF1B.GLD='SALES RETURN' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
    
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD,dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     
   
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='1' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group1=1)"
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='2' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group2=1)"
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group3=1)"
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='4' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group4=1)"

     
      rs5.Open "SELECT CREDITC.INVOICENO,CREDITC.INVOICEDATE AS D1 ,CREDITC.AMOUNT  FROM (INVOICEEND LEFT JOIN CREDITC ON INVOICEEND.TEXT = CREDITC.TEXT) LEFT JOIN CREDITA ON CREDITC.INVOICENO = CREDITA.INVOICENO where " & stringyear & " and CREDITC.AMOUNT >0  AND  iNVOICEEND.PrintOrder=70 AND CREDITA.district ='" + DWSalesReturn.Combosldistrictcode + "'", CON, adOpenDynamic, adLockOptimistic
      If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             DoEvents
             CON.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  where " & stringyear & " and  VNO= " & rs5!INVOICENO & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='I' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             
             rs5.MoveNext
             DoEvents
        Wend
     End If
     cr1.WindowTitle = "Dis.Wise Sales Return"
     cr1.ReportFileName = st1 & "\" & directory & "\report3.rpt"
     cr1.Action = 1
     Exit Sub
  End If
    
    cr1.WindowTitle = "Dis.Wise Sales Return"

    
  If Trim(Combosldistrictcode.Text) <> "" Then
      If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
            MsgBox "invalid date"
            Exit Sub
      End If
     Balance = 0
        
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C',CREDITA.SUBLEDGER as subleger ,CREDITA.agentname, CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) " & _
     "where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AgentName= '" & AgCombo.Text & "' and  CREDITA.district='" + DWSalesReturn.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER, CREDITA.AGENTNAME,CREDITA.district, BOOKS.GROUPCODE"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C',CREDITA.SUBLEDGER as subleger , CREDITA.agentname,CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) " & _
     "where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AgentName= '" & AgCombo.Text & "' and  CREDITA.district='" + DWSalesReturn.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER, CREDITA.AGENTNAME,CREDITA.district, BOOKS.GROUPCODE"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C',CREDITA.SUBLEDGER as subleger , CREDITA.agentname,CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) " & _
     "where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AgentName= '" & AgCombo.Text & "' and  CREDITA.district='" + DWSalesReturn.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER, CREDITA.AGENTNAME,CREDITA.district, BOOKS.GROUPCODE"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C',CREDITA.SUBLEDGER as subleger , CREDITA.agentname,CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) " & _
     "where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AgentName= '" & AgCombo.Text & "' and  CREDITA.district='" + DWSalesReturn.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER, CREDITA.AGENTNAME,CREDITA.district, BOOKS.GROUPCODE"

    
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND CNF1B.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND CNF1B.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND CNF1B.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND CNF1B.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"

     
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD,dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWSalesReturn.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     
    
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='1' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group1=1)"
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='2' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group2=1)"
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group3=1)"
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='4' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group4=1)"
     
     
     
     If rs5.State = 1 Then rs5.close
     rs5.Open "SELECT CREDITC.INVOICENO,CREDITC.INVOICEDATE AS D1 ,CREDITC.AMOUNT  FROM (INVOICEEND LEFT JOIN CREDITC ON INVOICEEND.TEXT = CREDITC.TEXT) LEFT JOIN CREDITA ON CREDITC.INVOICENO = CREDITA.INVOICENO where " & stringyear & " and CREDITC.AMOUNT >0  AND  iNVOICEEND.PrintOrder=70 AND CREDITA.district ='" + DWSalesReturn.Combosldistrictcode + "' AND CREDITA.AgentName='" & AgCombo.Text & "'", CON, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             CON.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  where " & stringyear & " and  VNO= " & rs5!INVOICENO & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='C' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             rs5.MoveNext
        Wend
     End If
   
     
     
     DoEvents
     DoEvents
     main.reportname = "Dis. Sales"
     ' viewlDisSales.genreport
     ' PrintOption.Show
  Else
      If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
         MsgBox "invalid date"
         Exit Sub
      End If

      
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C',CREDITA.SUBLEDGER as subleger ,CREDITA.AGENTNAME, CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AgentName= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER,CREDITA.AGENTNAME, CREDITA.district, BOOKS.GROUPCODE"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C',CREDITA.SUBLEDGER as subleger ,CREDITA.AGENTNAME, CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AgentName= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER,CREDITA.AGENTNAME, CREDITA.district, BOOKS.GROUPCODE"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C',CREDITA.SUBLEDGER as subleger ,CREDITA.AGENTNAME, CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AgentName= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER,CREDITA.AGENTNAME, CREDITA.district, BOOKS.GROUPCODE"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid) SELECT DISTINCT CREDITA.INVOICEDATE AS VDATE, CREDITA.INVOICENO AS VNO,'C',CREDITA.SUBLEDGER as subleger ,CREDITA.AGENTNAME, CREDITA.district AS DISTNAME, sum(CREDITB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid FROM (CREDITA LEFT JOIN (CREDITB LEFT JOIN BOOKS ON CREDITB.BOOKCODE = BOOKS.BOOKCODE) ON CREDITA.INVOICENO = CREDITB.INVOICENO) where " & stringyear & " and convert(smalldatetime,CREDITA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CREDITA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CREDITA.AgentName= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) GROUP BY CREDITA.INVOICEDATE, CREDITA.INVOICENO, CREDITA.SUBLEDGER,CREDITA.AGENTNAME, CREDITA.district, BOOKS.GROUPCODE"
      
      
      
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE, sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND CNF1B.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "'  GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE, sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND CNF1B.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE, sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND CNF1B.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE, sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND CNF1B.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"
     
     
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     CON.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND DNFB.GLD='SALES RETURN' and agentname='" & AgCombo.Text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
      
      
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='1' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group1=1)"
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='2' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group2=1)"
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group3=1)"
     CON.Execute "Update RPTTEMPINDIS1 set groupcheck='4' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group4=1)"
      
      
      
     rs5.Open "SELECT CREDITC.INVOICENO,CREDITC.INVOICEDATE AS D1 ,CREDITC.AMOUNT  FROM (CREDITEND LEFT JOIN CREDITC ON CREDITEND.TEXT = CREDITC.TEXT) LEFT JOIN CREDITA ON CREDITC.INVOICENO = CREDITA.INVOICENO where " & stringyear & " and CREDITC.AMOUNT >0  AND  CREDITEND.PrintOrder=70 AND  CREDITA.AgentName='" & AgCombo.Text & "'", CON, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             CON.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  where " & stringyear & " and  VNO= " & rs5!INVOICENO & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='I' and groupcheck='3'"
             rs5.MoveNext
        Wend
     End If
     
     
      
 Balance = 0
 main.reportname = "Dis. Sales"
  
 End If
  
  Dim r1 As New ADODB.Recordset
  Dim r2 As New ADODB.Recordset
  
  cr1.Reset
  cr1.ReportFileName = rptPath & "\report3.rpt"
  If RS.State = 1 Then RS.close
  RS.Open "select sum(BNETAMT) from RPTTEMPINDIS1", CON
  If RS.EOF = False Then
     cr1.Formulas(0) = "gtotal=" & RS(0) & ""
  End If
  cr1.WindowState = crptMaximized
  cr1.Action = 1


  Commandshow.Enabled = True
  Exit Sub


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

On Error GoTo aa2

Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> (Trim(UCase("MainMenu") Or Trim(UCase("frmbilllist"))))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop

aa2:


Me.Top = 0
Me.Left = 0
Set RS = New ADODB.Recordset
    CNSetup
    RS.Open "setup1", CON, adOpenDynamic, adLockReadOnly, adCmdTable
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
    RS.close
    RS.Open "select * from DISTRICTS order by DISTRICTNAME", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS!DISTRICTNAME
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    RS.Open "select Agentname  from AgentMaster where " & stringyear & " and " & stringyear & " order by AgentNAME", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
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
    RS.close



BackColorFrom Me

End Sub

