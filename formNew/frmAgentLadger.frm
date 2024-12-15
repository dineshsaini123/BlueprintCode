VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAgentLadger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Representative  Ladger"
   ClientHeight    =   6648
   ClientLeft      =   1680
   ClientTop       =   1392
   ClientWidth     =   8256
   Icon            =   "frmAgentLadger.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6648
   ScaleWidth      =   8256
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboGp 
      Height          =   315
      ItemData        =   "frmAgentLadger.frx":000C
      Left            =   4725
      List            =   "frmAgentLadger.frx":0019
      TabIndex        =   24
      Top             =   4545
      Width           =   795
   End
   Begin VB.CommandButton cmdRepAmtNew 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Export to Excel (Representative Wise Net Books Amt)"
      Height          =   555
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4500
      Width           =   2520
   End
   Begin VB.CommandButton cmdRepQtyNew 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Export to Excel (Representative Wise Net Books Qty)"
      Height          =   555
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3960
      Width           =   2520
   End
   Begin VB.ComboBox cbogd 
      Height          =   315
      Left            =   4620
      TabIndex        =   18
      Top             =   5220
      Width           =   795
   End
   Begin VB.CommandButton cmdGodownIssue 
      Caption         =   "&Challan Wise && Godown Issue"
      Height          =   555
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5100
      Width           =   2520
   End
   Begin MSComCtl2.DTPicker dateAsOn 
      Height          =   375
      Left            =   5805
      TabIndex        =   16
      Top             =   4020
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   656
      _Version        =   393216
      Format          =   142606337
      CurrentDate     =   42409
   End
   Begin VB.CommandButton cmdIssueAson 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Export to Excel (Representative Wise Books Issued as on)"
      Height          =   195
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Export to Excel (Representative Wise Net Books Qty && Net Amt)"
      Height          =   435
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.CheckBox Check1_selectAll 
      Caption         =   "Select All"
      Height          =   375
      Left            =   4860
      TabIndex        =   13
      Top             =   180
      Width           =   1035
   End
   Begin VB.CommandButton cmdamountbal 
      Caption         =   "&Balance Amount Group Wise"
      Height          =   495
      Left            =   5820
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.OptionButton OptionRec 
      Caption         =   "Receive"
      Height          =   255
      Left            =   6030
      TabIndex        =   11
      Top             =   5355
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.OptionButton OptionIssue 
      Caption         =   "Issue"
      Height          =   195
      Left            =   5820
      TabIndex        =   10
      Top             =   5340
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.ComboBox cboStation 
      Height          =   315
      Left            =   6120
      TabIndex        =   8
      Top             =   5220
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Representative Receive  Report"
      Height          =   555
      Left            =   1260
      TabIndex        =   7
      Top             =   3315
      Width           =   2520
   End
   Begin VB.CommandButton cmdsum 
      Caption         =   "&Joint Report"
      Height          =   495
      Left            =   6525
      TabIndex        =   6
      Top             =   5580
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.CommandButton cmdBilldet 
      Caption         =   "&Representative Issue Report"
      Height          =   555
      Left            =   1260
      TabIndex        =   5
      Top             =   2775
      Width           =   2520
   End
   Begin VB.ListBox cmbAgentName 
      Height          =   1344
      Left            =   1215
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   195
      Width           =   3555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print Representative Wise Balance Amount"
      Height          =   555
      Left            =   1260
      TabIndex        =   3
      Top             =   2205
      Width           =   2520
   End
   Begin Crystal.CrystalReport CR 
      Left            =   5490
      Top             =   5580
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   555
      Left            =   1260
      TabIndex        =   1
      Top             =   5700
      Width           =   2520
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Print Representative Wise Balance Books"
      Height          =   555
      Left            =   1260
      TabIndex        =   0
      Top             =   1650
      Width           =   2520
   End
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   375
      Left            =   3840
      TabIndex        =   22
      Top             =   4020
      Width           =   1665
      _ExtentX        =   2942
      _ExtentY        =   656
      _Version        =   393216
      Format          =   142606337
      CurrentDate     =   42409
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code"
      Height          =   315
      Left            =   3825
      TabIndex        =   25
      Top             =   4560
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   1185
      Left            =   1170
      Top             =   3915
      Width           =   6315
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   315
      Left            =   5550
      TabIndex        =   23
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Godown"
      Height          =   315
      Left            =   3900
      TabIndex        =   19
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Station By"
      Height          =   255
      Left            =   5820
      TabIndex        =   9
      Top             =   5100
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Representative :"
      Height          =   270
      Left            =   75
      TabIndex        =   2
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "frmAgentLadger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim s
Dim g1, g2, g3, g4 As String
Private Sub Check1_selectAll_Click()
For J = 0 To cmbAgentName.ListCount - 1
    cmbAgentName.Selected(J) = False
Next
If Check1_selectAll.value = 1 Then
  For J = 0 To cmbAgentName.ListCount - 1
      cmbAgentName.Selected(J) = True
  Next
End If
End Sub
Private Sub cmdamountbal_Click()
'--------------------------------------------
  Screen.MousePointer = vbHourglass
  Dim rs2 As New ADODB.Recordset
    '=====================================
    s = ""
    s1_ = ""
    s11 = ""
    
    For I = 0 To cmbAgentName.ListCount - 1
    If cmbAgentName.Selected(I) = True Then
    If s = "" Then
       s = "{INVOICEA.AGENTNAME}='" & cmbAgentName.List(I) & "'"
       s1_ = "a.AGENTNAME='" & cmbAgentName.List(I) & "'"
       s11 = "invoicea.AGENTNAME='" & cmbAgentName.List(I) & "'"
    Else
       s = s & " Or " & "{INVOICEA.AGENTNAME}='" & cmbAgentName.List(I) & "'"
       s1_ = s1_ & " Or " & "a.AGENTNAME='" & cmbAgentName.List(I) & "'"
       s11 = s11 & " Or " & "invoicea.AGENTNAME='" & cmbAgentName.List(I) & "'"
    End If
    End If
    Next
    
    
   Dim k1 As Integer
    
   
    
    If RS.State = 1 Then RS.close
    If s1_ <> "" Then
        RS.Open "select b.invoiceno,b.bookcode from INVOICEB_sp as b inner join INVOICEA_sp as a on a.invoiceno = b.invoiceno where a.fyear='" & session & "' and a.setupid='" & setupid & "' and " & s1_, con
    Else
        RS.Open "select b.invoiceno,b.bookcode from INVOICEB_sp as b inner join INVOICEA_sp as a on a.invoiceno = b.invoiceno where a.fyear='" & session & "' and a.setupid='" & setupid & "'", con
    End If
    
    
   While RS.EOF = False
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT BOOKS.groupcode FROM INVOICEB_sp LEFT JOIN BOOKS ON INVOICEB_sp.BOOKCODE = BOOKS.BOOKCODE where INVOICEB_sp.fyear='" & session & "' and INVOICEB_sp.setupid='" & setupid & "' and INVOICEB_sp.invoiceno=" & RS.Fields("invoiceno").value & " and INVOICEB_sp.bookcode='" & RS.Fields("bookcode").value & "'", con
    If rs1.EOF = False Then
       con.Execute "update INVOICEB_sp set group1=0,group2=0,group3=0,group4=0 where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
       con.Execute "update INVOICEB_sp set groupName='" & rs1(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
    End If
    
    '===================================
    
      If rs2.State = 1 Then rs2.close
      rs2.Open "SELECT NETAMOUNT,invoiceno,bookcode,groupName FROM INVOICEB_sp where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and  bookCODE='" & RS.Fields("bookcode").value & "'", con
      If rs2.EOF = False Then
         k1 = returnGroup(rs2!GroupName)
      If k1 = 1 Then
         con.Execute "update INVOICEB_sp set group1='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 2 Then
         con.Execute "update INVOICEB_sp set group2='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 3 Then
         con.Execute "update INVOICEB_sp set group3='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 4 Then
         con.Execute "update INVOICEB_sp set group4='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      End If

      End If
    
    RS.MoveNext
   Wend


'--For Return--------------------------------------------------------------------------------------


    k1 = 0
    
    If RS.State = 1 Then RS.close
    If s1_ <> "" Then
        RS.Open "select b.invoiceno,b.bookcode from INVOICEb_spret as b inner join INVOICEA_spret as a on a.invoiceno = b.invoiceno where a.fyear='" & session & "' and a.setupid='" & setupid & "' and " & s1_, con
    Else
        RS.Open "select b.invoiceno,b.bookcode from INVOICEb_spret as b inner join INVOICEA_spret as a on a.invoiceno = b.invoiceno where a.fyear='" & session & "' and a.setupid='" & setupid & "'", con
    End If
    
    While RS.EOF = False
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT BOOKS.groupcode FROM INVOICEb_spret LEFT JOIN BOOKS ON INVOICEb_spret.BOOKCODE = BOOKS.BOOKCODE where INVOICEb_spret.fyear='" & session & "' and INVOICEb_spret.setupid='" & setupid & "' and INVOICEb_spret.invoiceno=" & RS.Fields("invoiceno").value & " and INVOICEb_spret.bookcode='" & RS.Fields("bookcode").value & "'", con
    If rs1.EOF = False Then
       
       con.Execute "update INVOICEb_spret set group1=0,group2=0,group3=0,group4=0 where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
       con.Execute "update INVOICEb_spret set groupName='" & rs1(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
       
    End If
    
    '================================
    
      If rs2.State = 1 Then rs2.close
      rs2.Open "SELECT NETAMOUNT,invoiceno,bookcode,groupName FROM INVOICEb_spret where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and  bookCODE='" & RS.Fields("bookcode").value & "'", con
      If rs2.EOF = False Then
         k1 = returnGroup(rs2!GroupName)
      If k1 = 1 Then
         con.Execute "update INVOICEb_spret set group1='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 2 Then
         con.Execute "update INVOICEb_spret set group2='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 3 Then
         con.Execute "update INVOICEb_spret set group3='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 4 Then
         con.Execute "update INVOICEb_spret set group4='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      End If

      End If

    RS.MoveNext
    Wend

'==================================================================

Screen.MousePointer = vbDefault

'===================================================================











End Sub

Private Sub cmdBilldet_Click()
      
      If check_selectRep = False Then
       MsgBox "Select at least 1 Representative... ", vbCritical
       Exit Sub
    End If
    
  Dim date_ As String
  date_ = "(convert(smalldatetime,b.INVOICEDATE,103)>= convert(smalldatetime,'" & txtFrom.value & "',103) and convert(smalldatetime,b.INVOICEDATE,103)<= convert(smalldatetime,'" & dateAson.value & "',103)) "
  
  Screen.MousePointer = vbHourglass
  
  Dim rs2 As New ADODB.Recordset
    '=====================================
    s = ""
    s1__ = ""
    s11 = ""
    For I = 0 To cmbAgentName.ListCount - 1
    If cmbAgentName.Selected(I) = True Then
    If s = "" Then
       s = "{INVOICEA.AGENTNAME}='" & cmbAgentName.List(I) & "'"
       s1_ = "a.AGENTNAME='" & cmbAgentName.List(I) & "'"
       s11 = "INVOICEA.AGENTNAME='" & cmbAgentName.List(I) & "'"
    Else
       s = s & " Or " & "{INVOICEA.AGENTNAME}='" & cmbAgentName.List(I) & "'"
       s1_ = s1_ & " Or " & "a.AGENTNAME='" & cmbAgentName.List(I) & "'"
       s11 = s11 & " Or " & "INVOICEA.AGENTNAME='" & cmbAgentName.List(I) & "'"
    End If
    End If
    Next
    
    
   Dim k1 As Integer
    
   
    
    If RS.State = 1 Then RS.close
    If s1_ <> "" Then
        RS.Open "select b.invoiceno,b.bookcode from INVOICEB_sp as b inner join INVOICEA_sp as a on a.invoiceno = b.invoiceno where " & s1_, con
    Else
        RS.Open "select b.invoiceno,b.bookcode from INVOICEB_sp as b inner join INVOICEA_sp as a on a.invoiceno = b.invoiceno", con
    End If
    
    
   While RS.EOF = False
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT BOOKS.groupcode FROM INVOICEB_sp LEFT JOIN BOOKS ON INVOICEB_sp.BOOKCODE = BOOKS.BOOKCODE where INVOICEB_sp.fyear='" & session & "' and INVOICEB_sp.setupid='" & setupid & "' and  INVOICEB_sp.invoiceno=" & RS.Fields("invoiceno").value & " and INVOICEB_sp.bookcode='" & RS.Fields("bookcode").value & "'", con
    If rs1.EOF = False Then
       con.Execute "update INVOICEB_sp set group1=0,group2=0,group3=0,group4=0 where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
       con.Execute "update INVOICEB_sp set groupName='" & rs1(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
    End If
    
    '===================================
    
      If rs2.State = 1 Then rs2.close
      rs2.Open "SELECT NETAMOUNT,invoiceno,bookcode,groupName FROM INVOICEB_sp where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and  bookCODE='" & RS.Fields("bookcode").value & "'", con
      If rs2.EOF = False Then
         k1 = returnGroup(rs2!GroupName)
      If k1 = 1 Then
         con.Execute "update INVOICEB_sp set group1='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 2 Then
         con.Execute "update INVOICEB_sp set group2='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 3 Then
         con.Execute "update INVOICEB_sp set group3='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 4 Then
         con.Execute "update INVOICEB_sp set group4='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      End If

      End If
    
    RS.MoveNext
   Wend
    
     
    
    
    DSNNew
    
    If MsgBox("Want to View ?", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
    End If
    
    
    s = "(" & s & ")"
    
    cr.Reset
    cr.ReportFileName = rptPath & "/agentbill.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If s <> "" Then
        If cboStation.Text <> "" Then
           cr.ReplaceSelectionFormula s & " and " & "{INVOICEA.STATION}='" & cboStation.Text & "'"
           Else
           cr.ReplaceSelectionFormula "(" & s & " and " & "{INVOICEA.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {INVOICEA.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "'))"
        End If
    Else
        If cboStation.Text <> "" Then
           cr.ReplaceSelectionFormula "{INVOICEA.STATION}='" & cboStation.Text & "'"
           cr.Formulas(0) = "station1='" & cboStation.Text & "'"
        End If
    End If
    
    cr.Formulas(1) = "g1='" & g1 & "'"
    cr.Formulas(2) = "g2='" & g2 & "'"
    cr.Formulas(3) = "g3='" & g3 & "'"
    cr.Formulas(4) = "g4='" & g4 & "'"
    
    
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowSearchBtn = True

    
    cr.WindowState = crptMaximized
    cr.Action = 1
    
    
    Screen.MousePointer = vbDefault

End Sub
Public Function returnGroup(Str As String) As Integer
       
       If rss.State = 1 Then rss.close
       rss.Open "select * from groups where " & stringyear & " and groupcode='" & Str & "'", con
       If rss.EOF = False Then
        
        If rss!group1 = True Then
            returnGroup = 1
        ElseIf rss!group2 = True Then
            returnGroup = 2
        ElseIf rss!group3 = True Then
            returnGroup = 3
        ElseIf rss!group4 = True Then
            returnGroup = 4
        End If
        
       End If
              
End Function

Private Sub cmdexit_Click()
 Unload Me
End Sub
Sub querystring()
    s = ""
    
    For I = 0 To cmbAgentName.ListCount - 1
    
    If cmbAgentName.Selected(I) = True Then
    
        If s = "" Then
           s = "{ISSUEBOOK.AGENTNAME}='" & cmbAgentName.List(I) & "'"
        Else
           s = s & " Or " & "{ISSUEBOOK.AGENTNAME}='" & cmbAgentName.List(I) & "'"
        End If
    
    End If
    
    Next
    
End Sub
Function check_selectRep() As Boolean
check_selectRep = False
For J = 0 To cmbAgentName.ListCount - 1
    If cmbAgentName.Selected(J) = True Then
       check_selectRep = True
       Exit Function
    End If
Next

End Function

Private Sub cmdGodownIssue_Click()
 
DSNNew
    
cr.Reset
cr.ReportFileName = rptPath & "/PartyWiseItemWiseSpIssueDet.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
If cbogd.Text <> "" Then
   cr.ReplaceSelectionFormula "{PartyWiseItemWiseQty.godown}='" & cbogd.Text & "' and {PartyWiseItemWiseQty.INVOICEDATE}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.INVOICEDATE}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
Else
   cr.ReplaceSelectionFormula "{PartyWiseItemWiseQty.INVOICEDATE}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.INVOICEDATE}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
End If
cr.WindowShowPrintBtn = True
cr.WindowShowPrintSetupBtn = True
cr.WindowShowSearchBtn = True
cr.WindowState = crptMaximized
cr.Action = 1

End Sub

Private Sub cmdIssueAson_Click()
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String

On Error GoTo err:


If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Integer

c = 1
r = 1



xl.Columns("A:H").ColumnWidth = 18
J = 2

For I = 0 To cmbAgentName.ListCount - 1
  
    If cmbAgentName.Selected(I) = True Then
       r = 1
       
       xlSheet.Cells(r, J).value = cmbAgentName.List(I)
    
    ''Raws fill==========================================================
    Q1 = 0
    q2 = 0
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT BOOKCODE,BOOKNAME FROM invoiceSPBQry  group by BOOKCODE,BOOKNAME", con
    While RS.EOF = False
       
       Q1 = 0
       q2 = 0
       
       If rs_1.State = 1 Then rs_1.close
       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([QUANTITY]) as qty FROM invoiceSPBQry " & _
       " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and INVOICEDATE<=convert(smalldatetime,'" & dateAson.value & "',103)  group by BOOKCODE,agentname,BOOKNAME", con
       If rs_1.RecordCount > 0 Then
        If Not IsNull(rs_1(3)) Then
           Q1 = rs_1(3)
        End If
       End If
       
       'If rs_1.State = 1 Then rs_1.close
       'rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([QUANTITY]) as qty FROM invoiceSPRETBQry " & _
       '" where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "'  group by BOOKCODE,agentname,BOOKNAME", CON
       
       'If rs_1.RecordCount > 0 Then
       ' If Not IsNull(rs_1(3)) Then
       '    q2 = rs_1(3)
       ' End If
       'End If
       
       'Q1 = (Q1 - q2)
       
       
       r = r + 1
       
       xlSheet.Cells(r, 1).value = RS!Bookname
       xlSheet.Cells(r, J).value = Q1
       
       
      RS.MoveNext
    
    Wend
    
    '====================================================================
    
    J = J + 1
    End If

Next

Screen.MousePointer = vbDefault


Exit Sub

Screen.MousePointer = vbDefault

err:

MsgBox err.DESCRIPTION


End Sub

Private Sub cmdRepAmtNew_Click()
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_, date_ As String

On Error GoTo err:


If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Integer
Dim amt1, amt2 As Double

date_ = "(convert(smalldatetime,INVOICEDATE,103)>= convert(smalldatetime,'" & txtFrom.value & "',103) and convert(smalldatetime,INVOICEDATE,103)<= convert(smalldatetime,'" & dateAson.value & "',103)) "

c = 1
r = 1


If RS.State = 1 Then RS.close
If cbogp.Text = "" Then
   RS.Open "SELECT BOOKCODE,BOOKNAME FROM invoiceSPBQry  group by BOOKCODE,BOOKNAME", con
Else
   RS.Open "SELECT BOOKCODE,BOOKNAME FROM invoiceSPBQry where groupcode='" & cbogp & "'  group by BOOKCODE,BOOKNAME", con
End If


xl.Columns("A:H").ColumnWidth = 18
J = 3

For I = 0 To cmbAgentName.ListCount - 1
  
    If cmbAgentName.Selected(I) = True Then
       r = 1
       
       xlSheet.Cells(r, J).value = cmbAgentName.List(I)
       RS.MoveFirst
    
    ''Raws fill==========================================================
    Q1 = 0
    q2 = 0
    amt1 = 0
    amt2 = 0
    
    While RS.EOF = False
       
       Q1 = 0
       q2 = 0
       
       amt1 = 0
       amt2 = 0
       
       '======fatch Qty ====================================
       

       If rs_1.State = 1 Then rs_1.close
       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM invoiceSPBQry " & _
       " where (agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and " & date_ & ")  group by BOOKCODE,agentname,BOOKNAME", con
       If rs_1.RecordCount > 0 Then
        If Not IsNull(rs_1(3)) Then
           amt1 = rs_1(3)
        End If
       End If
       
       If rs_1.State = 1 Then rs_1.close
       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM invoiceSPRETBQry " & _
       " where (agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and " & date_ & ")  group by BOOKCODE,agentname,BOOKNAME", con
       
       If rs_1.RecordCount > 0 Then
        If Not IsNull(rs_1(3)) Then
           amt2 = rs_1(3)
        End If
       End If
       
       
       
       '======end code ====================================
       
       
       amt1 = (amt1 - amt2)
       r = r + 1
       xlSheet.Cells(r, 1).value = RS!Bookname
       xlSheet.Cells(r, 2).value = RS!Bookcode
       xlSheet.Cells(r, J).value = amt1
       
       
       
'       '======fatch Net ====================================
'
'       amt1 = 0
'       amt2 = 0
'
'
'       If rs_1.State = 1 Then rs_1.close
'       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM invoiceSPBQry " & _
'       " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "'  group by BOOKCODE,agentname,BOOKNAME", con
'       If rs_1.RecordCount > 0 Then
'        If Not IsNull(rs_1(3)) Then
'           amt1 = rs_1(3)
'        End If
'       End If
'
'       If rs_1.State = 1 Then rs_1.close
'       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM invoiceSPRETBQry " & _
'       " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "'  group by BOOKCODE,agentname,BOOKNAME", con
'
'       If rs_1.RecordCount > 0 Then
'        If Not IsNull(rs_1(3)) Then
'           amt2 = rs_1(3)
'        End If
'       End If
'
'       amt1 = (amt1 - amt2)
'       xlSheet.Cells(r, J + 1).value = Round(amt1, 0)
'       '======end code ====================================
       
       
       
       
      RS.MoveNext
    
    Wend
    
    '====================================================================
    
    J = J + 1
    End If

Next

Screen.MousePointer = vbDefault


Exit Sub

Screen.MousePointer = vbDefault

err:

MsgBox err.DESCRIPTION




End Sub

Private Sub cmdRepQty_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String

On Error GoTo err:


If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Integer
Dim amt1, amt2 As Double

c = 1
r = 1



xl.Columns("A:H").ColumnWidth = 18
J = 2

For I = 0 To cmbAgentName.ListCount - 1
  
    If cmbAgentName.Selected(I) = True Then
       r = 1
       
       xlSheet.Cells(r, J).value = cmbAgentName.List(I)
       
    
    ''Raws fill==========================================================
    Q1 = 0
    q2 = 0
    amt1 = 0
    amt2 = 0
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT BOOKCODE,BOOKNAME FROM invoiceSPBQry  group by BOOKCODE,BOOKNAME", con
    While RS.EOF = False
       
       Q1 = 0
       q2 = 0
       
       
       '======fatch Qty ====================================
       
       If rs_1.State = 1 Then rs_1.close
       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([QUANTITY]) as qty FROM invoiceSPBQry " & _
       " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "'  group by BOOKCODE,agentname,BOOKNAME", con
       If rs_1.RecordCount > 0 Then
        If Not IsNull(rs_1(3)) Then
           Q1 = rs_1(3)
        End If
       End If
       
       If rs_1.State = 1 Then rs_1.close
       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([QUANTITY]) as qty FROM invoiceSPRETBQry " & _
       " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "'  group by BOOKCODE,agentname,BOOKNAME", con
       
       If rs_1.RecordCount > 0 Then
        If Not IsNull(rs_1(3)) Then
           q2 = rs_1(3)
        End If
       End If
       
       '======end code ====================================
       
       
       Q1 = (Q1 - q2)
       r = r + 1
       xlSheet.Cells(r, 1).value = RS!Bookname
       xlSheet.Cells(r, J).value = Q1
       
       
       '======fatch Net ====================================
       
       amt1 = 0
       amt2 = 0

       
       If rs_1.State = 1 Then rs_1.close
       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM invoiceSPBQry " & _
       " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "'  group by BOOKCODE,agentname,BOOKNAME", con
       If rs_1.RecordCount > 0 Then
        If Not IsNull(rs_1(3)) Then
           amt1 = rs_1(3)
        End If
       End If
       
       If rs_1.State = 1 Then rs_1.close
       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM invoiceSPRETBQry " & _
       " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "'  group by BOOKCODE,agentname,BOOKNAME", con
       
       If rs_1.RecordCount > 0 Then
        If Not IsNull(rs_1(3)) Then
           amt2 = rs_1(3)
        End If
       End If
       
       amt1 = (amt1 - amt2)
       xlSheet.Cells(r, J + 1).value = Round(amt1, 0)
       '======end code ====================================
       
       
       
       
      RS.MoveNext
    
    Wend
    
    '====================================================================
    
    J = J + 2
    End If

Next

Screen.MousePointer = vbDefault


Exit Sub

Screen.MousePointer = vbDefault

err:

MsgBox err.DESCRIPTION






End Sub

Private Sub cmdRepQtyNew_Click()
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String
Dim date_ As String
On Error GoTo err:

date_ = "(convert(smalldatetime,INVOICEDATE,103)>= convert(smalldatetime,'" & txtFrom.value & "',103) and convert(smalldatetime,INVOICEDATE,103)<= convert(smalldatetime,'" & dateAson.value & "',103)) "

If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Integer
Dim amt1, amt2 As Double

c = 1
r = 1


If RS.State = 1 Then RS.close
If cbogp.Text = "" Then
   RS.Open "SELECT BOOKCODE,BOOKNAME FROM invoiceSPBQry  group by BOOKCODE,BOOKNAME", con, adOpenDynamic, adLockReadOnly
Else
   RS.Open "SELECT BOOKCODE,BOOKNAME FROM invoiceSPBQry where GROUPCODE='" & cbogp & "'  group by BOOKCODE,BOOKNAME", con, adOpenDynamic, adLockReadOnly
End If
    


xl.Columns("A:H").ColumnWidth = 18
J = 3

For I = 0 To cmbAgentName.ListCount - 1
  
    If cmbAgentName.Selected(I) = True Then
       r = 1
       xlSheet.Cells(r, J).value = cmbAgentName.List(I)
    
    ''Raws fill==========================================================
    Q1 = 0
    q2 = 0
    amt1 = 0
    amt2 = 0
    

    RS.MoveFirst
    
    While RS.EOF = False
       
       Q1 = 0
       q2 = 0
       
       
       '======fatch Qty ====================================
       
       If rs_1.State = 1 Then rs_1.close
       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([QUANTITY]) as qty FROM invoiceSPBQry " & _
       " where (agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and " & date_ & ")  group by BOOKCODE,agentname,BOOKNAME", con
       If rs_1.RecordCount > 0 Then
        If Not IsNull(rs_1(3)) Then
           Q1 = rs_1(3)
        End If
       End If
       
       If rs_1.State = 1 Then rs_1.close
       rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([QUANTITY]) as qty FROM invoiceSPRETBQry " & _
       " where (agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and " & date_ & ")  group by BOOKCODE,agentname,BOOKNAME", con
       
       If rs_1.RecordCount > 0 Then
        If Not IsNull(rs_1(3)) Then
           q2 = rs_1(3)
        End If
       End If
       
       '======end code ====================================
       
       Q1 = (Q1 - q2)
       r = r + 1
       xlSheet.Cells(r, 1).value = RS!Bookname
       xlSheet.Cells(r, 2).value = RS!Bookcode
       
       xlSheet.Cells(r, J).value = Q1
       
       
       
       
       
      RS.MoveNext
    
    Wend
    
    '====================================================================
    
    J = J + 1
    End If

Next

Screen.MousePointer = vbDefault


Exit Sub

Screen.MousePointer = vbDefault

err:

MsgBox err.DESCRIPTION




End Sub

Private Sub cmdshow_Click()

    DSNNew

    If check_selectRep = False Then
       MsgBox "Select at least 1 Representative... ", vbCritical
       Exit Sub
    End If
   
    querystring
    
    
    If cbogp = "" Then
    s = s & " and {ISSUEBOOK.fyear}='" & session & "' and {ISSUEBOOK.setupid}=" & setupid & ""
    Else
    s = s & " and {ISSUEBOOK.groupcode}='" & cbogp & "' and {ISSUEBOOK.fyear}='" & session & "' and {ISSUEBOOK.setupid}=" & setupid & ""
    End If
    
    cr.Reset
    cr.ReportFileName = rptPath & "/BookLadger.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    'CR.ReplaceSelectionFormula "{ISSUEBOOK.AGENTNAME}='" & cmbAgentName.Text & "'"
    If s <> "" Then
     cr.ReplaceSelectionFormula s
    End If
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowSearchBtn = True
    cr.WindowState = crptMaximized
    
    cr.Action = 1
End Sub

Private Sub cmdsum_Click()
   querystring
    


End Sub

Private Sub Command1_Click()
    
    DSNNew
    
    If check_selectRep = False Then
       MsgBox "Select at least 1 Representative... ", vbCritical
       Exit Sub
    End If

    
    querystring
    
    s = s & " and {ISSUEBOOK.fyear}='" & session & "' and {ISSUEBOOK.setupid}=" & setupid & ""
    
    cr.Reset
    cr.ReportFileName = rptPath & "/BookAmountWise.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    'CR.ReplaceSelectionFormula "{ISSUEBOOK.AGENTNAME}='" & cmbAgentName.Text & "'"
    If s <> "" Then
     cr.ReplaceSelectionFormula s
    End If
    
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowSearchBtn = True

    cr.WindowState = crptMaximized
    cr.Action = 1

End Sub

Private Sub Command2_Click()

     If check_selectRep = False Then
       MsgBox "Select at least 1 Representative... ", vbCritical
       Exit Sub
    End If



  Screen.MousePointer = vbHourglass
  
  DSNNew
    

   Dim rs2 As New ADODB.Recordset
   Dim k1 As Integer


    '=====================================
    s = ""
    s1_ = ""
    s11 = ""
    For I = 0 To cmbAgentName.ListCount - 1
    If cmbAgentName.Selected(I) = True Then
    If s = "" Then
       s = "{CREDITA.AGENTNAME}='" & cmbAgentName.List(I) & "'"
       s1_ = "a.AGENTNAME='" & cmbAgentName.List(I) & "'"
       s11 = "CREDITa.AGENTNAME='" & cmbAgentName.List(I) & "'"
    Else
       s = s & " Or " & "{CREDITA.AGENTNAME}='" & cmbAgentName.List(I) & "'"
       s1_ = s1_ & " Or " & "a.AGENTNAME='" & cmbAgentName.List(I) & "'"
       s11 = s11 & " Or " & "CREDITa.AGENTNAME='" & cmbAgentName.List(I) & "'"
    End If
    End If
    Next
    
    
    
    
    If RS.State = 1 Then RS.close
    If s1_ <> "" Then
        RS.Open "select b.invoiceno,b.bookcode from INVOICEb_spRet as b inner join INVOICEA_spRet as a on a.invoiceno = b.invoiceno where " & s1_, con
    Else
        RS.Open "select b.invoiceno,b.bookcode from INVOICEb_spRet as b inner join INVOICEA_spRet as a on a.invoiceno = b.invoiceno where a.fyear='" & session & "' and a.setupid='" & setupid & "'", con
    End If
    
    While RS.EOF = False
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT BOOKS.groupcode FROM INVOICEb_spRet LEFT JOIN BOOKS ON INVOICEb_spRet.BOOKCODE = BOOKS.BOOKCODE where INVOICEb_spRet.fyear='" & session & "' and INVOICEb_spRet.setupid='" & setupid & "' and INVOICEb_spRet.invoiceno=" & RS.Fields("invoiceno").value & " and INVOICEb_spRet.bookcode='" & RS.Fields("bookcode").value & "'", con
    If rs1.EOF = False Then
       
       con.Execute "update INVOICEb_spRet set group1=0,group2=0,group3=0,group4=0 where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
       con.Execute "update INVOICEb_spRet set groupName='" & rs1(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
       
    End If
    
    '================================
    
      If rs2.State = 1 Then rs2.close
      rs2.Open "SELECT NETAMOUNT,invoiceno,bookcode,groupName FROM INVOICEb_spRet where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and  bookCODE='" & RS.Fields("bookcode").value & "'", con
      If rs2.EOF = False Then
         k1 = returnGroup(rs2!GroupName)
      If k1 = 1 Then
         con.Execute "update INVOICEb_spRet set group1='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 2 Then
         con.Execute "update INVOICEb_spRet set group2='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 3 Then
         con.Execute "update INVOICEb_spRet set group3='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      ElseIf k1 = 4 Then
         con.Execute "update INVOICEb_spRet set group4='" & rs2(0) & "' where " & stringyear & " and invoiceno=" & RS.Fields("invoiceno").value & " and bookcode='" & RS.Fields("bookcode").value & "'"
      End If

      End If

    
    
    
    
    RS.MoveNext
    Wend
    
    

'=========================================

    s = ""
    
    For I = 0 To cmbAgentName.ListCount - 1
    If cmbAgentName.Selected(I) = True Then
    If s = "" Then
       s = "{credita.AGENTNAME}='" & cmbAgentName.List(I) & "'"
    Else
       s = s & " Or " & "{credita.AGENTNAME}='" & cmbAgentName.List(I) & "'"
    End If
    
    End If
    Next

    
    
    If MsgBox("Want to View ?", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
    End If
    
    
   s = "(" & s & ")"
    
    
    cr.Reset
    cr.ReportFileName = rptPath & "/AgentIssue.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If s <> "" Then
      'Cr.ReplaceSelectionFormula s & "and" & "{credita.FREIGHT}='" & cboStation.Text & "'"
      cr.ReplaceSelectionFormula s & " and " & "({credita.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {credita.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "'))"
   End If
    
    
    If cboStation.Text <> "" Then
       cr.ReplaceSelectionFormula "{credita.FREIGHT}='" & cboStation.Text & "' & and " & "{credita.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {credita.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
       cr.Formulas(0) = "station1='" & cboStation.Text & "'"
    End If
    
    cr.Formulas(1) = "g1='" & g1 & "'"
    cr.Formulas(2) = "g2='" & g2 & "'"
    cr.Formulas(3) = "g3='" & g3 & "'"
    cr.Formulas(4) = "g4='" & g4 & "'"
    
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowSearchBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
    

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
     
     
     
     '*******Agent  combo fill
    If RS.State = 1 Then RS.close
    'RS.Open "select  Agentname from AgentMaster where " & stringyear & " order by agentname", CON, adOpenDynamic, adLockReadOnly, adCmdText
    RS.Open "select Rep as Representative from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
    cmbAgentName.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cmbAgentName.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    RS.close
   
   
   
    If RS.State = 1 Then RS.close
    'RS.Open "select  * from groupheading where " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
    RS.Open "select  * from groupheading", con, adOpenDynamic, adLockReadOnly, adCmdText
    If RS.EOF = False Then
       g1 = RS(0)
       g2 = RS(1)
       g3 = RS(2)
       g4 = RS(3)
    End If
   
   
    If RS.State = 1 Then RS.close
    RS.Open "select * from GodownMaster where len(Godwn)<=3 and " & stringyear & " order by id", con, adOpenForwardOnly, adLockReadOnly
    cbogd.Clear
    If Not RS.EOF Then
    Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cbogd.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
    End If

   
   
   'dateAsOn.value = Format(Date, "dd/MM/yyyy")
   
   BackColorFrom Me
   
   
   If RS.State = 1 Then RS.close
   RS.Open "select yarfrom,yarto from setup1", con
   If RS.EOF = False Then
      txtFrom.value = Format(RS!yarfrom, "dd/MM/yyyy")
      dateAson.value = Format(Date, "dd/MM/yyyy")
   End If
   
   
   Me.Top = 200
   Me.Left = 200
   
   
End Sub

Private Sub Form_Resize()

Me.Left = (Me.Width) / 2
Me.Top = ((Me.Height) / 2) - 900

End Sub

Private Sub OptionIssue_Click()
    
If OptionIssue.value = True Then

    RS.Open "select  distinct(Station) from invoicea where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    cboStation.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cboStation.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    RS.close

 

 End If

End Sub

Private Sub OptionRec_Click()
 
 If OptionRec.value = True Then
    If RS.State = 1 Then RS.close
    RS.Open "select  distinct(FREIGHT) from credita where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
    cboStation.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cboStation.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    RS.close
    
 End If
 
    
End Sub
