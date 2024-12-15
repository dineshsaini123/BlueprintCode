VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGPSales 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6660
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6660
   Begin VB.ComboBox cbogp 
      Height          =   315
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   675
      Width           =   3885
   End
   Begin VB.ComboBox Combosldistrictcode 
      Height          =   315
      Left            =   6720
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2220
      Width           =   1425
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   630
      Left            =   2115
      Picture         =   "frmGPSales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1860
      Width           =   1545
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   630
      Left            =   3855
      Picture         =   "frmGPSales.frx":0BE4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1860
      Width           =   1455
   End
   Begin VB.ComboBox AgCombo 
      Height          =   288
      Left            =   2085
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   195
      Width           =   3885
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   360
      Top             =   1755
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
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   2070
      TabIndex        =   3
      Top             =   1155
      Width           =   1155
      _ExtentX        =   2032
      _ExtentY        =   614
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   3960
      TabIndex        =   5
      Top             =   1155
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   614
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Name"
      Height          =   195
      Left            =   270
      TabIndex        =   9
      Top             =   225
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1245
      Width           =   1065
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   315
      Left            =   3420
      TabIndex        =   7
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   735
      Width           =   615
   End
End
Attribute VB_Name = "frmGPSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RS As Recordset
Private Sub COMBOGENLEDGER_Change()
End Sub
Private Sub COMBOGENLEDGER_LostFocus()
COMBOGENLEDGER = UCase(COMBOGENLEDGER)
    If Trim(COMBOGENLEDGER.text) <> "" Then
        RS.Open "select * from gledger where " & stringyear & " and slf=FALSE", con, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.BOF Then
                COMBOGENLEDGER.SetFocus
        End If
        RS.close
    End If
End Sub

Private Sub Combosubledger_GotFocus()
    If Trim(COMBOGENLEDGER.text) = "" Then
        COMBOGENLEDGER.SetFocus
    End If
End Sub

Private Sub Combosubledger_LostFocus()
If Trim(Combosubledger.text) <> "" Then
    If Trim(COMBOGENLEDGER.text) <> "" Then
        If RS.State = 1 Then
            RS.close
        End If
        RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.close
    Else
        Combosubledger.text = ""
    End If
End If
End Sub

Private Sub Combosldistrictcode_KeyPress(KeyAscii As Integer)

If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   sendkeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
      sendkeys "{Down}"
      sendkeys "{tab}"
End If



End Sub

Private Sub Command1_Click()
End Sub

Private Sub CommandReturn_Click()
    ''MainMenu.Toolbar1.Visible = True
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
  
Screen.MousePointer = vbHourglass

  con.Execute "Delete  from rpttempindis1 where  len(SUBLEGER)>0"
  
  'CON.Execute "Delete  from rpttempindis1"
  
  If trs9.State = 1 Then trs9.close
  trs9.Open "Select *  from groups where " & stringyear & " and group1=1 order BY groupcode", con, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g1 = "( " & trs9!groupcode & " )"
    End If
    If trs9.State = 1 Then trs9.close
    trs9.Open "Select *  from groups where " & stringyear & " and group2=1 order BY groupcode", con, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g2 = "( " & trs9!groupcode & " )"
    End If
    
    If trs9.State = 1 Then trs9.close
    trs9.Open "Select *  from groups where " & stringyear & " and group3=1 order BY groupcode", con, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g3 = "( " & trs9!groupcode & " )"
    End If
    
    If trs9.State = 1 Then trs9.close
    trs9.Open "Select *  from groups where " & stringyear & " and group4=1 order BY groupcode", con, adOpenStatic, adLockOptimistic
    If trs9.RecordCount > 0 Then
       g4 = "( " & trs9!groupcode & " )"
    End If

    
    If trs.State = 1 Then trs.close
    con.Execute "Delete from treport"
    trs.Open "select * from treport where " & stringyear, con, adOpenDynamic, adLockOptimistic
    xstr = Me.date1 + "  To  " + Me.date2
    trs.AddNew
    trs!text = xstr
    trs!Genledger = g1
    trs!subledger = g2
    trs!narration = g3
    trs!header = g4
    
    
    trs.update
  If AgCombo.text = "" And Trim(Combosldistrictcode.text) <> "" Then
      If DateDiff("d", Trim(date1.text), Trim(date2.text)) <= 0 Then
            MsgBox "invalid date"
            Exit Sub
      End If
     Balance = 0
        
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.agentname, INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND   INVOICEA.district='" + Me.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group1=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND   INVOICEA.district='" + Me.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group2=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND   INVOICEA.district='" + Me.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group3=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND   INVOICEA.district='" + Me.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group4=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
    
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,-sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND CNF1B.GLD='SALES' and  SLEDGER.DISTCODE='" + Me.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND CNF1B.GLD='SALES' and SLEDGER.DISTCODE='" + Me.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND CNF1B.GLD='SALES' and SLEDGER.DISTCODE='" + Me.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND CNF1B.GLD='SALES' and SLEDGER.DISTCODE='" + Me.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + Me.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD,dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) and  CASHA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger ,CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  and  CASHA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME ,CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) and  CASHA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME ,CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) and  CASHA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"

     
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='1' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group1=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='2' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group2=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group3=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group4=1)"

     
     rs5.Open "SELECT INVOICEC.INVOICENO,INVOICEC.INVOICEDATE AS D1 ,INVOICEC.AMOUNT  FROM (INVOICEEND LEFT JOIN INVOICEC ON INVOICEEND.TEXT = INVOICEC.TEXT) LEFT JOIN INVOICEA ON INVOICEC.INVOICENO = INVOICEA.INVOICENO WHERE " & stringyear & " and INVOICEC.AMOUNT >0  AND  iNVOICEEND.PrintOrder=70 AND INVOICEA.district ='" + frmGPSales.Combosldistrictcode + "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             DoEvents
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  " & stringyear & " and VNO= " & rs5!invoiceNo & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='I' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             
             rs5.MoveNext
             DoEvents
        Wend
     End If
     If rs5.State = 1 Then rs5.close
     rs5.Open "SELECT CASHC.INVOICENO,CASHC.INVOICEDATE AS D1 ,CASHC.AMOUNT  FROM (CASHEND LEFT JOIN CASHC ON CASHEND.TEXT = CASHC.TEXT) LEFT JOIN CASHA ON CASHC.INVOICENO = CASHA.INVOICENO WHERE " & stringyear & " and CASHC.AMOUNT >0  AND  CASHEND.PrintOrder=70 AND CASHA.district ='" + frmGPSales.Combosldistrictcode + "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             DoEvents
             'If rs5!INVOICENO = 2594 Then MsgBox rs5!INVOICENO
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  " & stringyear & " and VNO= " & rs5!invoiceNo & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='S' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             rs5.MoveNext
        Wend
     End If
  
  
    Cr1.ReportFileName = st1 & "\" & directory & "\groupWiseSale.rpt"
    Cr1.Action = 1
    Exit Sub
  
  
  End If
    
    
    
  If Trim(Combosldistrictcode.text) <> "" Then
     If DateDiff("d", Trim(date1.text), Trim(date2.text)) <= 0 Then
            MsgBox "invalid date"
            Exit Sub
     End If
     Balance = 0
        
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.agentname, INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) " & _
     "WHERE " & stringyear & " and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.text & "' and  INVOICEA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where  " & stringyear & " and group1=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) " & _
     " WHERE " & stringyear & " and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.text & "' and  INVOICEA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO)" & _
     " WHERE " & stringyear & " and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.text & "' and  INVOICEA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger , INVOICEA.agentname,INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) " & _
     "WHERE " & stringyear & " and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.text & "' and  INVOICEA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER, INVOICEA.AGENTNAME,INVOICEA.district, BOOKS.GROUPCODE"

    
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,-sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,  CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT CNF1A.CND, CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD,CNF1A.AGENTNAME,  SLEDGER.DISTCODE, CNF1B.groupcode"
     
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group1=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD,dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DNFA.DND, DNFA.DNN,'D' as Vtype,  DNFA.PSLD, dnfa.agentname,SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid FROM SLEDGER RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where " & stringyear & " and GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' and SLEDGER.DISTCODE='" + frmGPSales.Combosldistrictcode + "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
    
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND CASHA.AgentName= '" & AgCombo.text & "' and  CASHA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group1=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger ,CASHA.AGENTNAME, CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND CASHA.AgentName= '" & AgCombo.text & "' and  CASHA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group2=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME ,CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND CASHA.AgentName= '" & AgCombo.text & "' and  CASHA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group3=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER,AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid) SELECT DISTINCT CASHA.INVOICEDATE AS VDATE,CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME ,CASHA.district AS DISTNAME, sum(CASHB.NETAMOUNT) AS BNETAMT,BOOKS.GROUPCODE," & UId & " as userid FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE " & stringyear & " and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND CASHA.AgentName= '" & AgCombo.text & "' and  CASHA.district='" + frmGPSales.Combosldistrictcode + "'and BOOKS.GROUPCODE In (select groupcode from groups where group4=1) GROUP BY CASHA.INVOICEDATE, CASHA.INVOICENO, CASHA.SUBLEDGER, CASHA.AGENTNAME,CASHA.district, BOOKS.GROUPCODE"
     
     
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='1' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group1=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='2' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group2=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group3=1)"
     con.Execute "Update RPTTEMPINDIS1 set groupcheck='4' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group4=1)"
     
     
     
     If rs5.State = 1 Then rs5.close
     rs5.Open "SELECT INVOICEC.INVOICENO,INVOICEC.INVOICEDATE AS D1 ,INVOICEC.AMOUNT  FROM (INVOICEEND LEFT JOIN INVOICEC ON INVOICEEND.TEXT = INVOICEC.TEXT) LEFT JOIN INVOICEA ON INVOICEC.INVOICENO = INVOICEA.INVOICENO WHERE " & stringyear & " and INVOICEC.AMOUNT >0  AND  iNVOICEEND.PrintOrder=70 AND INVOICEA.district ='" + frmGPSales.Combosldistrictcode + "' AND INVOICEA.AgentName='" & AgCombo.text & "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  " & stringyear & " and  VNO= " & rs5!invoiceNo & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='I' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             rs5.MoveNext
        Wend
     End If
     If rs5.State = 1 Then rs5.close
     rs5.Open "SELECT CASHC.INVOICENO,CASHC.INVOICEDATE AS D1 ,CASHC.AMOUNT  FROM (CASHEND LEFT JOIN CASHC ON CASHEND.TEXT = CASHC.TEXT) LEFT JOIN CASHA ON CASHC.INVOICENO = CASHA.INVOICENO WHERE " & stringyear & " and CASHC.AMOUNT >0  AND  CASHEND.PrintOrder=70 AND CASHA.district ='" + frmGPSales.Combosldistrictcode + "' AND CASHA.AgentName='" & AgCombo.text & "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  " & stringyear & " and VNO= " & rs5!invoiceNo & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='S' and groupcheck='3'"
             For I = 1 To 10000
               DoEvents
             Next I
             rs5.MoveNext
        Wend
     End If
     
     
     
     
     DoEvents
     DoEvents
     main.reportname = "Dis. Sales"
  Else
      If DateDiff("d", Trim(date1.text), Trim(date2.text)) <= 0 Then
         MsgBox "invalid date"
         Exit Sub
      End If
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid,setupid,fyear) " & _
      "SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.AGENTNAME," & _
      " INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid," & setupid & ",'" & session & "' FROM " & _
      "(INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) " & _
      " WHERE INVOICEA.fyear='" & session & "' and INVOICEA.setupid='" & setupid & "' and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And " & _
      " convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND INVOICEA.AgentName= '" & AgCombo.text & "' and BOOKS.GROUPCODE " & _
      " In (select groupcode from groups where " & stringyear & " and group1=1) GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER,INVOICEA.AGENTNAME, INVOICEA.district, BOOKS.GROUPCODE"
      
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid,setupid,fyear) " & _
      "SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.AGENTNAME," & _
      " INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid," & setupid & ",'" & session & "' FROM " & _
      "(INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) WHERE  INVOICEA.fyear='" & session & "' and INVOICEA.setupid='" & setupid & "' and " & _
      " convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND " & _
      "INVOICEA.AgentName= '" & AgCombo.text & "' and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group2=1) " & _
      " GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER,INVOICEA.AGENTNAME, INVOICEA.district, BOOKS.GROUPCODE"
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid,setupid,fyear) " & _
      "SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.AGENTNAME, " & _
      " INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid," & setupid & ",'" & session & "' FROM " & _
      " (INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) " & _
      " WHERE INVOICEA.fyear='" & session & "' and INVOICEA.setupid='" & setupid & "' and convert(smalldatetime,INVOICEA.INVOICEDATE,193) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND " & _
      " INVOICEA.AgentName= '" & AgCombo.text & "' and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group3=1) " & _
      " GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER,INVOICEA.AGENTNAME, INVOICEA.district, BOOKS.GROUPCODE"
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT,  GROUPCODE,userid,setupid,fyear) " & _
      "SELECT DISTINCT INVOICEA.INVOICEDATE AS VDATE, INVOICEA.INVOICENO AS VNO,'I',INVOICEA.SUBLEDGER as subleger ,INVOICEA.AGENTNAME," & _
      " INVOICEA.district AS DISTNAME, sum(INVOICEB.NETAMOUNT) AS BNETAMT, BOOKS.GROUPCODE, " & UId & " as userid," & setupid & ",'" & session & "' FROM " & _
      "(INVOICEA LEFT JOIN (INVOICEB LEFT JOIN BOOKS ON INVOICEB.BOOKCODE = BOOKS.BOOKCODE) ON INVOICEA.INVOICENO = INVOICEB.INVOICENO) " & _
      " WHERE INVOICEA.fyear='" & session & "' and INVOICEA.setupid='" & setupid & "' and convert(smalldatetime,INVOICEA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And convert(smalldatetime,INVOICEA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " & _
      " AND INVOICEA.AgentName= '" & AgCombo.text & "' and BOOKS.GROUPCODE In (select groupcode from groups where " & stringyear & " and group4=1)" & _
      " GROUP BY INVOICEA.INVOICEDATE, INVOICEA.INVOICENO, INVOICEA.SUBLEDGER,INVOICEA.AGENTNAME, INVOICEA.district, BOOKS.GROUPCODE"
      
       
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid,setupid,fyear) SELECT CNF1A.CND, " & _
     " CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid ," & setupid & ",'" & session & "'" & _
     " FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where CNF1A.fyear='" & session & "' and CNF1A.setupid='" & setupid & "' and groupcode " & _
     " In (select groupcode from groups where " & stringyear & " and group1=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.text & "'  GROUP BY CNF1A.CND, " & _
     " CNF1A.CNN,CNF1A.PSLD,cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid,setupid,fyear) SELECT CNF1A.CND, CNF1A.CNN,'C'" & _
     " as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid," & setupid & ",'" & session & "' FROM SLEDGER RIGHT " & _
     " JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where CNF1A.fyear='" & session & "' and CNF1A.setupid='" & setupid & "' and GROUPCODE In (select groupcode from groups " & _
     " where " & stringyear & " and group2=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.text & "' GROUP BY CNF1A.CND, CNF1A.CNN,CNF1A.PSLD, cnf1a.agentname, " & _
     " SLEDGER.DISTCODE, CNF1B.groupcode"
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid,setupid,fyear) SELECT CNF1A.CND," & _
     " CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid," & setupid & ",'" & session & "' " & _
     " FROM SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where CNF1A.fyear='" & session & "' and CNF1A.setupid='" & setupid & "' and GROUPCODE In " & _
     "(select groupcode from groups where " & stringyear & " and group3=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.text & "' GROUP BY CNF1A.CND, CNF1A.CNN," & _
     "CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE,userid,setupid,fyear) SELECT CNF1A.CND, " & _
     "CNF1A.CNN,'C' as Vtype,  CNF1A.PSLD,CNF1A.AGENTNAME, SLEDGER.DISTCODE,- sum( CNF1B.A) as bnetamt, CNF1B.groupcode," & UId & " as userid," & setupid & ",'" & session & "'  FROM" & _
     " SLEDGER RIGHT JOIN (CNF1B RIGHT JOIN CNF1A ON CNF1B.CNN = CNF1A.CNN) ON SLEDGER.SUBLEDGER = CNF1A.PSLD where CNF1A.fyear='" & session & "' and CNF1A.setupid='" & setupid & "' and GROUPCODE In " & _
     "(select groupcode from groups where " & stringyear & " and group4=1) AND CNF1B.GLD='SALES' and agentname='" & AgCombo.text & "' GROUP BY CNF1A.CND, " & _
     " CNF1A.CNN,CNF1A.PSLD, cnf1a.agentname, SLEDGER.DISTCODE, CNF1B.groupcode"

     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid,setupid,fyear) SELECT DNFA.DND, DNFA.DNN,'D'" & _
     " as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid," & setupid & ",'" & session & "' FROM SLEDGER " & _
     " RIGHT JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where DNFA.fyear='" & session & "' and DNFA.setupid='" & setupid & "' and GROUPCODE In (select groupcode from groups" & _
     " where " & stringyear & " and group1=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, " & _
     "SLEDGER.DISTCODE,  DNFB.groupcode"
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid,setupid,fyear) SELECT DNFA.DND, DNFA.DNN,'D'" & _
     " as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid," & setupid & ",'" & session & "' FROM SLEDGER RIGHT" & _
     " JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where DNFA.fyear='" & session & "' and DNFA.setupid='" & setupid & "' and GROUPCODE In (select groupcode from groups" & _
     " where " & stringyear & " and group2=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname," & _
     " SLEDGER.DISTCODE,  DNFB.groupcode"
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid,setupid,fyear) SELECT DNFA.DND, DNFA.DNN,'D'" & _
     " as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid," & setupid & ",'" & session & "' FROM SLEDGER RIGHT JOIN" & _
     " (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where DNFA.fyear='" & session & "' and DNFA.setupid='" & setupid & "' and GROUPCODE In (select groupcode from groups where " & stringyear & " and " & _
     " group3=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname, SLEDGER.DISTCODE,  DNFB.groupcode"
     
     con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME,DISTNAME, BNETAMT, GROUPCODE,userid,setupid,fyear) SELECT DNFA.DND, DNFA.DNN," & _
     "'D' as Vtype,  DNFA.PSLD, DNFA.AGENTNAME, SLEDGER.DISTCODE,sum( DNFB.A) as bnetamt, DNFB.groupcode," & UId & " as userid," & setupid & ",'" & session & "' FROM SLEDGER RIGHT" & _
     " JOIN (DNFB RIGHT JOIN DNFA ON DNFB.DNN = DNFA.DNN) ON SLEDGER.SUBLEDGER = DNFA.PSLD where DNFA.fyear='" & session & "' and DNFA.setupid='" & setupid & "' and GROUPCODE In (select groupcode from groups" & _
     " where " & stringyear & " and group4=1) AND DNFB.GLD='SALES' and agentname='" & AgCombo.text & "' GROUP BY DNFA.DND, DNFA.dNN,DNFA.PSLD, dnfa.agentname," & _
     " SLEDGER.DISTCODE,  DNFB.groupcode"


      
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE, userid,setupid,fyear) SELECT DISTINCT " & _
      "CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME, " & _
      "sum(CASHB.NETAMOUNT) AS BNETAMT,  BOOKS.GROUPCODE , " & UId & " as userid," & setupid & ",'" & session & "' FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON " & _
      "CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE casha.fyear='" & session & "' and casha.setupid='" & setupid & "' and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And " & _
      " convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND CASHA.AGENTNAME= '" & AgCombo.text & "' and BOOKS.GROUPCODE In " & _
      " (select groupcode from groups where " & stringyear & " and group1=1) GROUP BY cashA.INVOICEDATE, cashA.INVOICENO, cashA.SUBLEDGER, CASHA.AGENTNAME,cashA.district, BOOKS.GROUPCODE"
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE, userid,setupid,fyear) SELECT DISTINCT " & _
      " CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME," & _
      " sum(CASHB.NETAMOUNT) AS BNETAMT,  BOOKS.GROUPCODE , " & UId & " as userid," & setupid & ",'" & session & "' FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON " & _
      "CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE casha.fyear='" & session & "' and casha.setupid='" & setupid & "' and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) And " & _
      " convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND CASHA.AGENTNAME= '" & AgCombo.text & "' and BOOKS.GROUPCODE In " & _
      "(select groupcode from groups where " & stringyear & " and group2=1) GROUP BY cashA.INVOICEDATE, cashA.INVOICENO, cashA.SUBLEDGER, CASHA.AGENTNAME," & _
      "cashA.district, BOOKS.GROUPCODE"
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE, userid,setupid,fyear) SELECT DISTINCT " & _
      " CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME," & _
      " sum(CASHB.NETAMOUNT) AS BNETAMT,  BOOKS.GROUPCODE , " & UId & " as userid," & setupid & ",'" & session & "' FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON " & _
      "CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE casha.fyear='" & session & "' and casha.setupid='" & setupid & "' and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) " & _
      " And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND CASHA.AGENTNAME= '" & AgCombo.text & "' and BOOKS.GROUPCODE In " & _
      " (select groupcode from groups where " & stringyear & " and group3=1) GROUP BY cashA.INVOICEDATE, cashA.INVOICENO, cashA.SUBLEDGER, CASHA.AGENTNAME," & _
      "cashA.district, BOOKS.GROUPCODE"
      
      con.Execute "insert into RPTTEMPINDIS1(vdate,vno,vtype,SUBLEGER, AGENTNAME, DISTNAME, BNETAMT, GROUPCODE, userid,setupid,fyear) SELECT DISTINCT " & _
      "CASHA.INVOICEDATE AS VDATE, CASHA.INVOICENO AS VNO,'S',CASHA.SUBLEDGER as subleger , CASHA.AGENTNAME, CASHA.district AS DISTNAME," & _
      " sum(CASHB.NETAMOUNT) AS BNETAMT,  BOOKS.GROUPCODE , " & UId & " as userid," & setupid & ",'" & session & "' FROM (CASHA LEFT JOIN (CASHB LEFT JOIN BOOKS ON " & _
      "CASHB.BOOKCODE = BOOKS.BOOKCODE) ON CASHA.INVOICENO = CASHB.INVOICENO) WHERE casha.fyear='" & session & "' and casha.setupid='" & setupid & "' and convert(smalldatetime,CASHA.INVOICEDATE,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) " & _
      " And convert(smalldatetime,CASHA.INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND CASHA.AGENTNAME= '" & AgCombo.text & "' and BOOKS.GROUPCODE In " & _
      "(select groupcode from groups where " & stringyear & " and group4=1) GROUP BY cashA.INVOICEDATE, cashA.INVOICENO, cashA.SUBLEDGER, CASHA.AGENTNAME," & _
      "cashA.district, BOOKS.GROUPCODE"

      
      con.Execute "Update RPTTEMPINDIS1 set groupcheck='1' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group1=1)"
      con.Execute "Update RPTTEMPINDIS1 set groupcheck='2' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group2=1)"
      con.Execute "Update RPTTEMPINDIS1 set groupcheck='3' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group3=1)"
      con.Execute "Update RPTTEMPINDIS1 set groupcheck='4' where " & stringyear & " and groupcode In (select groupcode from groups where " & stringyear & " and group4=1)"

      
      
     rs5.Open "SELECT INVOICEC.INVOICENO,INVOICEC.INVOICEDATE AS D1 ,INVOICEC.AMOUNT  FROM (INVOICEEND LEFT JOIN INVOICEC ON " & _
     "INVOICEEND.TEXT = INVOICEC.TEXT) LEFT JOIN INVOICEA ON INVOICEC.INVOICENO = INVOICEA.INVOICENO WHERE INVOICEA.setupid=" & setupid & " and INVOICEA.fyear='" & session & "'  and INVOICEC.AMOUNT >0  AND " & _
     " iNVOICEEND.PrintOrder=70 AND  INVOICEA.AgentName='" & AgCombo.text & "'", con, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
            
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  " & stringyear & " and VNO= " & rs5!invoiceNo & " AND convert(smalldatetime,VDATE,103)= convert(smalldatetime,'" & rs5!d1 & "',103) AND VTYPE='I' and groupcheck='3'"
             rs5.MoveNext
        Wend
     End If
     If rs5.State = 1 Then rs5.close
     rs5.Open "SELECT CASHC.INVOICENO,CASHC.INVOICEDATE AS D1 ,CASHC.AMOUNT  FROM (invoiceEnd LEFT JOIN CASHC ON invoiceEnd.TEXT = CASHC.TEXT) LEFT JOIN CASHA ON CASHC.INVOICENO = CASHA.INVOICENO WHERE CASHC.AMOUNT >0  AND  invoiceEnd.PrintOrder=70 AND invoiceEnd.type='cash' and  CASHA.AgentName='" & AgCombo.text & "'", con, adOpenDynamic, adLockOptimistic
     '''rs5.Open "SELECT CASHC.INVOICENO,CASHC.INVOICEDATE AS D1 ,CASHC.AMOUNT  FROM (CASHEND LEFT JOIN CASHC ON CASHEND.TEXT = CASHC.TEXT) LEFT JOIN CASHA ON CASHC.INVOICENO = CASHA.INVOICENO WHERE CASHC.AMOUNT >0  AND  CASHEND.PrintOrder=70 AND  CASHA.AgentName='" & AgCombo.Text & "'", CON, adOpenDynamic, adLockOptimistic
     If rs5.RecordCount > 0 Then
         While Not rs5.EOF
             con.Execute "Update RPTTEMPINDIS1 set BNETAMT= BNETAMT-" & rs5!amount & "  WHERE  " & stringyear & " and VNO= " & rs5!invoiceNo & " AND VDATE= CDATE('" & rs5!d1 & "') AND VTYPE='S' and groupcheck='3'"
             For I = 1 To 1000
               DoEvents
             Next I
             rs5.MoveNext
        Wend
     End If
     
      
      Balance = 0
      main.reportname = "Dis. Sales"

End If


DSNNew

Dim r1 As New ADODB.Recordset
Dim r2 As New ADODB.Recordset
 
Cr1.Reset
Cr1.ReportFileName = rptPath & "\groupWiseSale.rpt"
Cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=dinesh.123;"

If cbogp.text <> "" Then
   Cr1.ReplaceSelectionFormula "{invoiceBQry.groupcode}='" & cbogp.text & "' and {invoiceBQry.AGENTNAME}='" & AgCombo.text & "'"

  If (UCase(cbogp.text) = "LOW PRICE") Then
   Cr1.ReplaceSelectionFormula "{invoiceBQry.SERNAME}='" & cbogp.text & "' and {invoiceBQry.AGENTNAME}='" & AgCombo.text & "'"
  End If
Else
   Cr1.ReplaceSelectionFormula "{invoiceBQry.AGENTNAME}='" & AgCombo.text & "'"
End If
Cr1.WindowState = crptMaximized
Cr1.WindowShowPrintSetupBtn = True
Cr1.Action = 1
Commandshow.Enabled = True


Screen.MousePointer = vbDefault

End Sub
Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  sendkeys "{tab}"
End If
End Sub

Private Sub date1_LostFocus()
    If Trim(date1.text) <> "" Then
        If Not checkdate(Trim(date1.text), date1) Then
            date1.SetFocus
        End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  sendkeys "{tab}"
End If
End Sub

Private Sub date2_LostFocus()
    If Trim(date2.text) <> "" Then
        If Not checkdate(Trim(date2.text), date2) Then
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

Me.top = 0
Me.Left = 0
'Set CON = New ADODB.Connection
Set RS = New ADODB.Recordset
'    With CON
 '       .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
 '       .Open
 '   End With
    CNSetup
    RS.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly
    date1.text = RS!yarfrom
    date2.text = RS!yarto
    RS.close
    
    'RS.Open "select * from DISTRICTS order by DISTRICTNAME", con, adOpenDynamic, adLockReadOnly, adCmdText
    RS.Open "select * from groups where " & stringyear & " order by groupcode", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            'Me.Combosldistrictcode.AddItem RS(0)
            Me.cbogp.AddItem RS(0)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    
    
    Me.cbogp.AddItem "LOW PRICE"
    
    Set RS = New ADODB.Recordset
    RS.Open "select Rep as Representative,Email from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
    Me.AgCombo.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            'Me.cmbAgentName.AddItem RS(0)
            Me.AgCombo.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    
    
    
'    RS.close
'    RS.Open "select  Agentname  from AgentMaster where " & stringyear & " order by AgentNAME", con, adOpenForwardOnly, adLockReadOnly, adCmdText
'    If Not RS.EOF Then
'        Do While Not RS.EOF
'              If IsNull(RS!agentname) = False Then
'                Me.AgCombo.AddItem RS!agentname
'            End If
'            If Not RS.EOF Then
'                RS.MoveNext
'            End If
'        Loop
'    End If
'    RS.close



BackColorFrom Me

End Sub


