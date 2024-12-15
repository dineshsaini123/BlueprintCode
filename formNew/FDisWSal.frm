VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form DWsales 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3432
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   7404
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3432
   ScaleWidth      =   7404
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport cr1 
      Left            =   495
      Top             =   2295
      _ExtentX        =   593
      _ExtentY        =   593
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
   Begin VB.ComboBox AgCombo 
      Height          =   315
      Left            =   2205
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   420
      Width           =   3885
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   720
      Left            =   3930
      Picture         =   "FDisWSal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2085
      Width           =   1455
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   720
      Left            =   2190
      Picture         =   "FDisWSal.frx":0BE4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2085
      Width           =   1455
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   2190
      TabIndex        =   2
      Top             =   1380
      Width           =   1155
      _ExtentX        =   2032
      _ExtentY        =   614
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox Combosldistrictcode 
      Height          =   315
      Left            =   2190
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   900
      Width           =   3885
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   4080
      TabIndex        =   3
      Top             =   1380
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   614
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "District Name"
      Height          =   240
      Left            =   375
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   315
      Left            =   3570
      TabIndex        =   8
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   1470
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Name"
      Height          =   195
      Left            =   390
      TabIndex        =   6
      Top             =   450
      Width           =   1650
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
End Sub
Private Sub COMBOGENLEDGER_LostFocus()
COMBOGENLEDGER = UCase(COMBOGENLEDGER)
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        RS.Open "select * from gledger where slf=FALSE", con, adOpenDynamic, adLockReadOnly, adCmdText
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
        RS.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.close
    Else
        Combosubledger.Text = ""
    End If
End If
End Sub

Private Sub AgCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim RS As New ADODB.Recordset
    RS.Open "select  * from DISTRICTS  where  AGENTNAME= '" + AgCombo.Text + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
        Combosldistrictcode.Clear
    If Not RS.EOF Then
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

Private Sub AgCombo_LostFocus()
  Dim RS As New ADODB.Recordset

 If AgCombo.Text <> "" Then
    RS.Open "select  * from DISTRICTS  where  AGENTNAME= '" + AgCombo.Text + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
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
  Else
  
     
     RS.Open "select  * from DISTRICTS", con, adOpenForwardOnly, adLockReadOnly, adCmdText
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

Private Sub Command1_Click()
End Sub

Private Sub CommandReturn_Click()
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
  con.Execute "Delete  from rpttempindis1"
  
  If trs9.State = 1 Then trs9.close
  trs9.Open "Select *  from groups where group1=1 order BY groupcode", con, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g1 = "( " & trs9!groupcode & " )"
    End If
    If trs9.State = 1 Then trs9.close
    trs9.Open "Select *  from groups where group2=1 order BY groupcode", con, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g2 = "( " & trs9!groupcode & " )"
    End If
    
    If trs9.State = 1 Then trs9.close
    trs9.Open "Select *  from groups where group3=1 order BY groupcode", con, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g3 = "( " & trs9!groupcode & " )"
    End If
    
    If trs9.State = 1 Then trs9.close
    trs9.Open "Select *  from groups where group4=1 order BY groupcode", con, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g4 = "( " & trs9!groupcode & " )"
    End If

    
    If trs.State = 1 Then trs.close
    con.Execute "Delete from treport"
    trs.Open "treport", con, adOpenDynamic, adLockOptimistic, adCmdTable
    xstr = DWsales.date1 + "  To  " + DWsales.date2
    trs.AddNew
    trs!Text = xstr
    trs!Genledger = g1
    trs!subledger = g2
    trs!narration = g3
    trs!header = g4
    
    
    trs.update
  If AgCombo.Text = "" And Trim(Combosldistrictcode.Text) <> "" Then
      If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
            MsgBox "invalid date"
            Exit Sub
      End If
     Balance = 0
        
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.agentname, INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND   INVOICEA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group1=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND   INVOICEA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group2=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND   INVOICEA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group3=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND   INVOICEA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group4=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
    
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,-sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group1=1) AND CNF1B.GLD='SALES' and  SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group2=1) AND CNF1B.GLD='SALES' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group3=1) AND CNF1B.GLD='SALES' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group4=1) AND CNF1B.GLD='SALES' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group1=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD,dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group2=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group3=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group4=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and  CASHA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group1=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger ,CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  and  CASHA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group2=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME ,CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and  CASHA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group3=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME ,CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and  CASHA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group4=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"

     
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='1' where groupcode In (select groupcode from groups where group1=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='2' where groupcode In (select groupcode from groups where group2=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where groupcode In (select groupcode from groups where group3=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where groupcode In (select groupcode from groups where group4=1)"

     
     rs5.Open "SELECT INVOICEC.INVOICENO,INVOICEC.INVOICEDATE AS D1 ,INVOICEC.AMOUNT  FROM (INVOICEEND LEFT JOIN INVOICEC ON INVOICEEND.TEXT = INVOICEC.TEXT) LEFT JOIN INVOICEA ON INVOICEC.INVOICENO = INVOICEA.INVOICENO WHERE INVOICEC.AMOUNT >0  AND  iNVOICEEND.PrintOrder=70 AND INVOICEA.district ='" + DWsales.Combosldistrictcode + "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             DoEvents
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  VNO= " & rs5!invoiceNo & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='I' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             
             rs5.MoveNext
             DoEvents
        Wend
     End If
     If rs5.State = 1 Then rs5.close
     rs5.Open "SELECT CASHC.INVOICENO,CASHC.INVOICEDATE AS D1 ,CASHC.AMOUNT  FROM (CASHEND LEFT JOIN CASHC ON CASHEND.TEXT = CASHC.TEXT) LEFT JOIN CASHA ON CASHC.INVOICENO = CASHA.INVOICENO WHERE CASHC.AMOUNT >0  AND  CASHEND.PrintOrder=70 AND CASHA.district ='" + DWsales.Combosldistrictcode + "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             DoEvents
             'If rs5!INVOICENO = 2594 Then MsgBox rs5!INVOICENO
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  VNO= " & rs5!invoiceNo & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='S' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             rs5.MoveNext
        Wend
     End If
  
  
    cr1.ReportFileName = rptPath & "\report2.rpt"
    cr1.Action = 1
    Exit Sub
  
  
  End If
    
    
    
  If Trim(Combosldistrictcode.Text) <> "" Then
     If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
            MsgBox "invalid date"
            Exit Sub
     End If
     Balance = 0
        
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.agentname, INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.Text & "' and  INVOICEA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group1=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.Text & "' and  INVOICEA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group2=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.Text & "' and  INVOICEA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group3=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.Text & "' and  INVOICEA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group4=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"

    
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,-sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group1=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group2=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group3=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group4=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group1=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD,dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group2=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group3=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group4=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' and SLEDGER.DISTCODE='" + DWsales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
    
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.AgentName= '" & AgCombo.Text & "' and  CASHA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group1=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger ,CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.AgentName= '" & AgCombo.Text & "' and  CASHA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group2=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME ,CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.AgentName= '" & AgCombo.Text & "' and  CASHA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group3=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME ,CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.AgentName= '" & AgCombo.Text & "' and  CASHA.district='" + DWsales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group4=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     
     
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='1' where groupcode In (select groupcode from groups where group1=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='2' where groupcode In (select groupcode from groups where group2=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where groupcode In (select groupcode from groups where group3=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='4' where groupcode In (select groupcode from groups where group4=1)"
     
     
     
     If rs5.State = 1 Then rs5.close
     rs5.Open "SELECT INVOICEC.INVOICENO,INVOICEC.INVOICEDATE AS D1 ,INVOICEC.AMOUNT  FROM (INVOICEEND LEFT JOIN INVOICEC ON INVOICEEND.TEXT = INVOICEC.TEXT) LEFT JOIN INVOICEA ON INVOICEC.INVOICENO = INVOICEA.INVOICENO WHERE INVOICEC.AMOUNT >0  AND  iNVOICEEND.PrintOrder=70 AND INVOICEA.district ='" + DWsales.Combosldistrictcode + "' AND INVOICEA.AgentName='" & AgCombo.Text & "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  VNO= " & rs5!invoiceNo & " AND convert(smalldatetime,VDATE,103)= convert(smalldatetime,'" & rs5!d1 & "',103) AND VTYPE='I' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             rs5.MoveNext
        Wend
     End If
     If rs5.State = 1 Then rs5.close
     rs5.Open "SELECT CASHC.INVOICENO,CASHC.INVOICEDATE AS D1 ,CASHC.AMOUNT  FROM (invoiceEND LEFT JOIN CASHC ON invoiceEND.TEXT = CASHC.TEXT) LEFT JOIN CASHA ON CASHC.INVOICENO = CASHA.INVOICENO WHERE CASHC.AMOUNT >0  AND  invoiceend.PrintOrder=70 and invoiceend.type='cash'  AND CASHA.district ='" + DWsales.Combosldistrictcode + "' AND CASHA.AgentName='" & AgCombo.Text & "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  VNO= " & rs5!invoiceNo & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='S' and groupcheck='3'"
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
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.AGENTNAME, INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where group1=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER,INVOICEA.AGENTNAME, INVOICEA.district, BOOKS.GROUPCODE"
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.AGENTNAME, INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where group2=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER,INVOICEA.AGENTNAME, INVOICEA.district, BOOKS.GROUPCODE"
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.AGENTNAME, INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where group3=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER,INVOICEA.AGENTNAME, INVOICEA.district, BOOKS.GROUPCODE"
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.AGENTNAME, INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where group4=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER,INVOICEA.AGENTNAME, INVOICEA.district, BOOKS.GROUPCODE"

       
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group1=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.Text & "'  GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group2=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.Text & "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group3=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.Text & "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where GROUPCODE In (select groupcode from groups where group4=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.Text & "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"

     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group1=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group2=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group3=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where GROUPCODE In (select groupcode from groups where group4=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.Text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"


      
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE, userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,  BOOKS.GROUPCODE , " & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.AGENTNAME= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where group1=1) GROUP BY cashA.INVOICEDATE, cashA.INVOICENO, cashA.SUBLEDGER, CASHA.AGENTNAME,cashA.district, BOOKS.GROUPCODE"
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE, userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,  BOOKS.GROUPCODE , " & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.AGENTNAME= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where group2=1) GROUP BY cashA.INVOICEDATE, cashA.INVOICENO, cashA.SUBLEDGER, CASHA.AGENTNAME,cashA.district, BOOKS.GROUPCODE"
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE, userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,  BOOKS.GROUPCODE , " & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.AGENTNAME= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where group3=1) GROUP BY cashA.INVOICEDATE, cashA.INVOICENO, cashA.SUBLEDGER, CASHA.AGENTNAME,cashA.district, BOOKS.GROUPCODE"
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE, userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,  BOOKS.GROUPCODE , " & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND CASHA.AGENTNAME= '" & AgCombo.Text & "' and BOOKS.GROUPCODE In (select groupcode from groups where group4=1) GROUP BY cashA.INVOICEDATE, cashA.INVOICENO, cashA.SUBLEDGER, CASHA.AGENTNAME,cashA.district, BOOKS.GROUPCODE"

      
      con.Execute "Update RPTTEMPINDIS1 set groupcheck='1' where groupcode In (select groupcode from groups where group1=1)"
      con.Execute "Update RPTTEMPINDIS1 set groupcheck='2' where groupcode In (select groupcode from groups where group2=1)"
      con.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where groupcode In (select groupcode from groups where group3=1)"
      con.Execute "Update RPTTEMPINDIS1 set groupcheck='4' where groupcode In (select groupcode from groups where group4=1)"

      
      
     rs5.Open "SELECT INVOICEC.INVOICENO,INVOICEC.INVOICEDATE AS D1 ,INVOICEC.AMOUNT  FROM (INVOICEEND LEFT JOIN INVOICEC ON INVOICEEND.TEXT = INVOICEC.TEXT) LEFT JOIN INVOICEA ON INVOICEC.INVOICENO = INVOICEA.INVOICENO WHERE INVOICEC.AMOUNT >0  AND  iNVOICEEND.PrintOrder=70 AND  INVOICEA.AgentName='" & AgCombo.Text & "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
            
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  VNO= " & rs5!invoiceNo & " AND convert(smalldatetime,VDATE,103)= convert(smalldatetime,'" & rs5!d1 & "',103) AND VTYPE='I' and groupcheck='3'"
             rs5.MoveNext
        Wend
     End If
     If rs5.State = 1 Then rs5.close
     rs5.Open "SELECT CASHC.INVOICENO,CASHC.INVOICEDATE AS D1 ,CASHC.AMOUNT  FROM (INVOICEEND LEFT JOIN CASHC ON INVOICEEND.TEXT = CASHC.TEXT) LEFT JOIN CASHA ON CASHC.INVOICENO = CASHA.INVOICENO WHERE CASHC.AMOUNT >0  AND   INVOICEEND.PrintOrder=70 AND  INVOICEEND.type='cash' and   CASHA.AgentName='" & AgCombo.Text & "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  VNO= " & rs5!invoiceNo & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='S' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             rs5.MoveNext
        Wend
     End If
     
      
 Balance = 0
 main.reportname = "Dis. Sales"
 
End If
 

Dim r1 As New ADODB.Recordset
Dim r2 As New ADODB.Recordset
's1 = App.Path & "\2003-04\tchitra.mdb"
'CON1.Execute "Delete * from rpttempindis1"
'CON.Execute "insert into  RPTTEMPINDIS1  IN '" & s1 & "' select * from RPTTEMPINDIS1"

'cr1.ReportFileName = App.Path & "\2003-04\report2.rpt"

DSNNew

cr1.Reset
cr1.ReportFileName = rptPath & "\report2.rpt"
cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=sidc;"
If RS.State = 1 Then RS.close
RS.Open "select sum(BNETAMT) from RPTTEMPINDIS1", con
If RS.EOF = False Then
 cr1.Formulas(0) = "gtotal=" & RS(0) & ""
End If
cr1.WindowState = crptMaximized
cr1.Action = 1

Commandshow.Enabled = True
con.Execute "Delete  from rpttempindis1"
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

On Error GoTo aa1

Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> (Trim(UCase("MainMenu") Or Trim(UCase("frmbilllist"))))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop


aa1:

Me.Top = 0
Me.Left = 0
'Set CON = New ADODB.Connection
Set RS = New ADODB.Recordset
'    With CON
 '       .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
 '       .Open
 '   End With
    CNSetup
    RS.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
    RS.close
    RS.Open "select * from DISTRICTS order by DISTRICTNAME", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS!DISTRICTNAME
          
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    RS.Open "select  Agentname  from AgentMaster where " & stringyear & " order by AgentNAME", con, adOpenForwardOnly, adLockReadOnly, adCmdText
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

