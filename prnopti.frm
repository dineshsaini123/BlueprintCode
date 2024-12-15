VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PrintOption 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Print Option"
   ClientHeight    =   1764
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5124
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1764
   ScaleWidth      =   5124
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Cdlbox 
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
   End
   Begin VB.CommandButton cmdprinter 
      Caption         =   "Send &Mail"
      Height          =   555
      Left            =   3360
      TabIndex        =   2
      Top             =   660
      Width           =   1305
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Screen"
      Height          =   555
      Left            =   1740
      TabIndex        =   1
      Top             =   660
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print View"
      Default         =   -1  'True
      Height          =   555
      Left            =   225
      TabIndex        =   0
      Top             =   660
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   180
      TabIndex        =   3
      Top             =   405
      Width           =   4836
      Begin VB.CommandButton Command4 
         Caption         =   "Send whatsapp"
         Height          =   555
         Left            =   4932
         TabIndex        =   6
         Top             =   252
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Order Dispatched"
         Height          =   555
         Left            =   4755
         TabIndex        =   4
         Top             =   1212
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1305
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox Check1_sms 
      Caption         =   "Send Mail With SMS"
      Height          =   375
      Left            =   3375
      TabIndex        =   5
      Top             =   1092
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "PrintOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprinter_Click()

If s1 = 0 Then
   MsgBox "Error...", vbCritical
   Exit Sub
End If

Dim rep_email, head_mail, mob As String
Dim M_Mailid As String

Dim rs_ As New ADODB.Recordset
Dim rs_mn As New ADODB.Recordset

Dim randomId As String
Dim s10 As String
mob = ""
randomId = ""
M_Mailid = ""


If rs_.State = 1 Then rs_.Close
rs_.Open "select count(*) from MailDetails", con
If rs_.RecordCount >= 100 Then
   MsgBox "Already 100 Mail Pending ...", vbCritical
   Exit Sub
End If

'''========================== update Random Id=============================
'If rs1.State = 1 Then rs1.close
'rs1.Open "select invoiceno,RandomNo FROM INVOICEA where RandomId=''", con
'While rs1.EOF = False
'
'If Len(rs1!RandomNo) >= 5 Then
'   s10 = Mid(value_, 1, 2)
'Else
'   s10 = Mid(value_, 1, 3)
'End If
'
's10 = "IN" & s10 & rs1!RandomNo
'
'con.Execute "update INVOICEA set Randomid= '" & s10 & "'  where invoiceno=" & rs1!invoiceNo & ""
'
'rs1.MoveNext
'Wend
'
'''=========================end code=======================================





If s1 = 1 Then
    
     
   '=====================================
    head_mail = ""
    If RS.State = 1 Then RS.Close
    RS.Open "select party,ADDRESS1,ADDRESS2,district,states,THROUGH,THROUGH1,BUNDLES,STATION,mobile from invoiceaQry where invoiceno=" & invoice.I_NO & "", con
    'RS.Open sql_, con
    If RS.EOF = False Then
       address1 = RS!party
       address2 = RS!address1
       address3 = RS!address2
       address4 = RS!District + ",(" + RS!states + ")"
       through = RS!through + " " + RS!through1
       bundle = RS!bundles & ""
       transport = RS!station & ""
       If Not IsNull(RS!mobile) Then
          mob = Mid(RS!mobile, 1, 10) & ""
       End If
       
      If rs_mn.State = 1 Then rs_mn.Close
        
       
    End If

    '=====================================
   If invoice.txtMark.text <> "NS" Then
   
            If Len(invoice.LblRandomNo) >= 5 Then
               s10 = Mid(value_, 1, 2)
            Else
               s10 = Mid(value_, 1, 3)
            End If
            
            s10 = "IN" & s10 & invoice.LblRandomNo
            
            con.Execute "update INVOICEA set smsDate=Convert(smalldatetime,'" & Date & "', 103),Randomid= '" & s10 & "',mobile='" & mob & "',SMSSend='n'  where invoiceno=" & invoice.I_NO & ""
     End If

    
   ''''----------------------------------------------------------
    
    
    If invoice.lblMail.Caption = "" Then
        MsgBox "Mail Can'nt Be Send..,Because Party Mail Id is not Found "
        Exit Sub
    End If
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "select Email,HeadEmail FROM Rep where rep='" & invoice.cmbAgentName & "'", CON_blue
    If RS.EOF = False Then
      
      If (RS!email <> "") Then
         rep_email = RS!email
      Else
         rep_email = "-"
      End If
      
      If Not IsNull(RS!HeadEmail) Then
         head_mail = RS!HeadEmail
      Else
         head_mail = "-"
      End If
         
    End If
    
    
    
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "select * from MailDetails where (Bill=" & invoice.I_NO & " and BillType='invoice')", con
    If RS.EOF = True Then
       con.Execute "insert into MailDetails(Bill,BillType,Mail,MailSended,OrderDis,address1,address2,address3,address4,through,Transport,Bundle,RepEmail,HeadEmail)" & _
       " values(" & invoice.I_NO & ",'invoice','" & invoice.lblMail & "','n','y','" & address1 & "','" & address2 & "','" & address3 & "','" & address4 & "','" & through & "','" & transport & "','" & bundle & "','" & rep_email & "','" & head_mail & "')"
    Else
       con.Execute "update MailDetails set MailSended='n',OrderDis='y',Mail='" & invoice.lblMail & "',RepEmail='" & rep_email & "',HeadEmail='" & head_mail & "',address1='" & address1 & "',address2='" & address2 & "',address3='" & address3 & "',address4='" & address4 & "' where (Bill=" & invoice.I_NO & " and BillType='invoice')"
    End If
    
    
    
    
    
    
    
    
ElseIf s1 = 2 Then
    
    head_mail = ""

    If RS.State = 1 Then RS.Close
    RS.Open "SELECT INVOICENO,MARKA,BUNDLES,THROUGH,STATION,THROUGH1,AgentName,Add1,Add2,district,City,States,mail,Randomno FROM INVOICEA_sp where invoiceno=" & frmBookIssueSp.I_NO.text & "", con
    If RS.EOF = False Then
       address1 = RS!agentname
       address2 = RS!add1
       address3 = RS!add2
       address4 = RS!District + ",(" + RS!states + ")"
       through = RS!through + " " + RS!through1
       bundle = RS!bundles & ""
       transport = RS!station & ""
       mail = RS!mail & ""
       randomId = RS!RandomNo
        
       If RS!agentname <> "" Then
       
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select Email,headEmail,Phone FROM Rep where rep='" & RS!agentname & "'", CON_blue
        If rs1.EOF = False Then
        If (rs1!HeadEmail <> "" Or Not IsNull(rs1!HeadEmail)) Then
           head_mail = rs1!HeadEmail
        Else
           head_mail = "-"
        End If
        
        mob = Mid(rs1!phone, 1, 10)
        
        End If

       End If
       
    End If
    
    ''===============================================
    '----SMS Code
    '===============================================
   
   If frmBookIssueSp.cboGodown.text <> "NS" Then

        'If Check1_sms.value = 1 Then

'           mob = "9997314681"

            If Len(randomId) >= 5 Then
               s10 = Mid(value_, 1, 2)
            Else
               s10 = Mid(value_, 1, 3)
            End If
            s10 = "SP" & s10 & randomId
         
            con.Execute "update INVOICEA_sp set smsDate=Convert(smalldatetime,'" & Date & "', 103),Randomid= '" & s10 & "',mobile='" & mob & "',SMSSend='n'  where invoiceno=" & frmBookIssueSp.I_NO & ""
            'smsSend frmBookIssueSp.I_NO, mob, s10
            
       ' End If

     End If

    
    ''''-----------------------------------------------------------------------------------
    ''===============================================
    
    If mail = "" Then
       MsgBox "Mail Can'nt Be Send..,Because Representative Mail Id is not Found "
       Exit Sub
    End If
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "select * from MailDetails where (Bill=" & frmBookIssueSp.I_NO & " and BillType='invoice_sp')", con
    If RS.EOF = True Then
       con.Execute "insert into MailDetails(Bill,BillType,Mail,MailSended,OrderDis,address1,address2,address3,address4,through,Transport,Bundle,HeadEmail)" & _
       " values(" & frmBookIssueSp.I_NO.text & ",'invoice_sp','" & mail & "','n','y','" & address1 & "','" & address2 & "','" & address3 & "','" & address4 & "','" & through & "','" & transport & "','" & bundle & "','" & head_mail & "')"
    Else
       con.Execute "update MailDetails set MailSended='n',OrderDis='y',Mail='" & mail & "',HeadEmail='" & head_mail & "',address1='" & address1 & "',address2='" & address2 & "',address3='" & address3 & "',address4='" & address4 & "' where (Bill=" & frmBookIssueSp.I_NO & " and BillType='invoice_sp')"
    End If

ElseIf s1 = 12 Then

    head_mail = ""
    
    If RS.State = 1 Then RS.Close
    RS.Open "select party,ADDRESS1,ADDRESS2,district,states,THROUGH,THROUGH1,BUNDLES,STATION from CREDITAQry where invoiceno=" & Critnote.I_NO & "", con
    If RS.EOF = False Then
    
       address1 = RS!party
       address2 = RS!address1
       address3 = RS!address2
       address4 = RS!District + ",(" + RS!states + ")"
       through = RS!through + " " + RS!through1
       bundle = RS!bundles & ""
       transport = RS!station & ""
    End If
    
    
    If Critnote.lblMail.Caption = "" Then
        MsgBox "Mail Can'nt Be Send..,Because Party Mail Id is not Found "
        Exit Sub
    End If
    
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "select Email,HeadEmail FROM Rep where rep='" & Critnote.cmbAgentName & "'", CON_blue
    If RS.EOF = False Then
      
      If (RS!email <> "") Then
         rep_email = RS!email
      Else
         rep_email = "-"
      End If
      
      If Not IsNull(RS!HeadEmail) Then
         head_mail = RS!HeadEmail
      Else
         head_mail = "-"
      End If
         
         
    End If
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "select * from MailDetails where (Bill=" & Critnote.I_NO & " and BillType='credit')", con
    If RS.EOF = True Then
       con.Execute "insert into MailDetails(Bill,BillType,Mail,MailSended,OrderDis,address1,address2,address3,address4,through,Transport,Bundle,RepEmail,HeadEmail)" & _
       " values(" & Critnote.I_NO.text & ",'credit','" & Critnote.lblMail & "','n','y','" & address1 & "','" & address2 & "','" & address3 & "','" & address4 & "','" & through & "','" & transport & "','" & bundle & "','" & rep_email & "','" & head_mail & "')"
    Else
       con.Execute "update MailDetails set MailSended='n',OrderDis='y',Mail='" & Critnote.lblMail & "',RepEmail='" & rep_email & "',HeadEmail='" & head_mail & "',address1='" & address1 & "',address2='" & address2 & "',address3='" & address3 & "',address4='" & address4 & "' where (Bill=" & Critnote.I_NO & " and BillType='credit')"
    End If




ElseIf s1 = 7 Then


Dim add1 As String
Dim partyname As String
Dim bill As Integer
Dim pmail As String


If SLEDGERPRINT.Combosubledger.text <> "" Then

con.Execute ("exec partyledger '" & setupid & "','" & session & "'" & _
",'" & SLEDGERPRINT.Combosubledger.text & "','" & SLEDGERPRINT.date1.text & "'" & _
",'" & SLEDGERPRINT.date2.text & "','" & UId & "','" & SLEDGERPRINT.COMBOGENLEDGER.text & "'")


add1 = ""
partyname = ""



If rs1.State = 1 Then rs1.Close
rs1.Open "select party,subledger,email from SLEDGER where subledger='" & SLEDGERPRINT.Combosubledger.text & "'", con
If rs1.EOF = False Then

   If InStr(rs1!party, "IMPREST") > 0 Then
      add1 = Mid(rs1!party, 1, Len(rs1!party) - 13)
   Else
      add1 = rs1!party
   End If
   
   partyname = rs1!subledger
   pmail = rs1!email
End If


If RS.State = 1 Then RS.Close
RS.Open "select max(bill) from MailDetails where billtype='ledger'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   If IsNull(RS(0)) Then
      bill = 9999 + 1
   Else
      bill = RS(0) + 1
   End If
  
End If

If Len(pmail) > 10 Then
   con.Execute "insert into MailDetails(Bill,Mail,BillType,rptName,address1,partyname) values('" & bill & "','" & pmail & "','ledger','SLLedger.rpt','" & add1 & "','" & partyname & "')"
End If

    





Else

For k5 = 0 To SLEDGERPRINT.cboacc_multiple.ListCount - 1

If SLEDGERPRINT.cboacc_multiple.Selected(k5) = True Then

    con.Execute ("exec partyledger '" & setupid & "','" & session & "'" & _
    ",'" & SLEDGERPRINT.cboacc_multiple.List(k5) & "','" & SLEDGERPRINT.date1.text & "'" & _
    ",'" & SLEDGERPRINT.date2.text & "','" & UId & "','" & SLEDGERPRINT.COMBOGENLEDGER.text & "'")
    
add1 = ""
partyname = ""



If rs1.State = 1 Then rs1.Close
rs1.Open "select party,subledger,email from SLEDGER where subledger='" & SLEDGERPRINT.cboacc_multiple.List(k5) & "'", con
If rs1.EOF = False Then

   If InStr(rs1!party, "IMPREST") > 0 Then
      add1 = Mid(rs1!party, 1, Len(rs1!party) - 13)
   Else
      add1 = rs1!party
   End If
   
   partyname = rs1!subledger
   pmail = rs1!email
End If


If RS.State = 1 Then RS.Close
RS.Open "select max(bill) from MailDetails where billtype='ledger'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   If IsNull(RS(0)) Then
      bill = 9999 + 1
   Else
      bill = RS(0) + 1
   End If
  
End If

If Len(pmail) > 10 Then
   con.Execute "insert into MailDetails(Bill,Mail,BillType,rptName,address1,partyname) values('" & bill & "','" & pmail & "','ledger','SLLedger.rpt','" & add1 & "','" & partyname & "')"
End If


End If



Next


End If
End If




MsgBox "Mail has been send...", vbInformation

End Sub
Private Sub Command1_Click()

Screen.MousePointer = vbHourglass

DSNNew

On Error Resume Next

Dim bill_, bill As String
bill_ = ""
bill = ""

If Len(invoice.I_NO) > 6 Then
    bill = Mid(Str(invoice.I_NO), 6)
    bill_ = billformat & "" & Format(Trim(bill), "00000")
 Else
    bill_ = billformat & "" & Format(Trim(Str(invoice.I_NO)), "00000")
 End If
 
DSNNew


If s1 = 501 Then



    cr1.Reset
    cr1.ReportFileName = rptPath & "/CASHBOOK.RPT"
    cr1.Connect = "filedsn=chitradsn;uid=" & sql_user & ";pwd=" & sql_pass
    cr1.WindowShowPrintSetupBtn = True
    cr1.WindowState = crptMaximized
    cr1.Formulas(1) = "fromDate='" & CASHBOOK.date1.text & "'"
    cr1.Formulas(2) = "toDate='" & CASHBOOK.date2.text & "'"
    cr1.Formulas(7) = "cname='" & UCase(cname_1) & "'"
    cr1.Formulas(8) = "add1='" & cname_2 & "'"
    cr1.Formulas(9) = "add2='" & cname_add1 & "'"
    cr1.Formulas(10) = "gst='" & gst & "'"
    cr1.Action = 1


ElseIf s1 = "1" Then
If PopUpValue6 = "withheader" Then
    cr1.Reset
    cr1.ReportFileName = rptPath & "/invoice_header.rpt"
    cr1.Connect = "filedsn=chitradsn;uid=" & sql_user & ";pwd=" & sql_pass
    cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & invoice.I_NO.text & ""
    amt = (invoice.txtAmtwords.text)
    cr1.Formulas(2) = "RSS='" & amt & "'"
    cr1.Formulas(3) = "tin='" & tin & "'"
    cr1.WindowShowPrintSetupBtn = True
    cr1.WindowState = crptMaximized
    cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
    cr1.Formulas(8) = "unit_='" & cname_2 & "'"
    cr1.Formulas(9) = "gst='" & gst & "'"
    cr1.Formulas(10) = "inv='" & bill_ & "'"
    cr1.Action = 1
ElseIf invoice.Check1_trans.value = 1 Then
        
'    sss1 = ""
'    sss2 = ""
'    sss3 = ""
'    h1 = 1
'
'    If RS.State = 1 Then RS.close
'    RS.Open "SELECT BOOKNAME,sum(QUANTITY-BQUANTITY) AS QTY FROM tmpSaleOrder1 WHERE INVOICENO='" & invoice.txtOrderNo.Text & "' group by BOOKNAME", con
'    While RS.EOF = False
'       If RS(1) > 0 Then
'
'          If h1 <= 3 Then
'             If sss1 = "" Then
'                sss1 = RS(0) & " : " & RS(1)
'              Else
'                sss1 = sss1 & " " & RS(0) & " : " & RS(1)
'             End If
'          End If
'
'          If (h1 > 4 And h1 <= 8) Then
'             If sss2 = "" Then
'                sss2 = RS(0) & " : " & RS(1)
'              Else
'                sss2 = sss2 & " " & RS(0) & " : " & RS(1)
'             End If
'          End If
'
'          If (h1 > 8 And h1 <= 12) Then
'             If sss3 = "" Then
'                sss3 = RS(0) & " : " & RS(1)
'              Else
'                sss3 = sss3 & " " & RS(0) & " : " & RS(1)
'             End If
'          End If
'
'
'
'          h1 = h1 + 1
'
'       End If
'       RS.MoveNext
'    Wend
'
'    If sss1 <> "" Then
'    ''sss1 = "Pending Books against your order. : - " & sss1 & " for above, please do not send the order again. The pending book will be sent shortly."
'      sss1 = "" & sss1
'    End If
'
    Dim d_ As Integer
    d_ = 1
    
    cr1.Reset
    If invoice.Check1_notPrint_inst.value = 0 Then
       cr1.ReportFileName = rptPath & "/invoice_Trans.rpt"
    Else
       cr1.ReportFileName = rptPath & "/invoice_Trans_image.rpt"
    End If
    cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & invoice.I_NO.text & ""
    amt = (invoice.txtAmtwords.text)
    cr1.Formulas(2) = "RSS='" & amt & "'"
    cr1.Formulas(3) = "tin='" & tin & "'"
    cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
    cr1.Formulas(8) = "unit_='" & cname_2 & "'"
    cr1.Formulas(9) = "inv='" & bill_ & "'"
    cr1.Formulas(10) = "gst='" & gst & "'"
    'cr1.Formulas(11) = "pending_books='" & sss1 & "'"
    'cr1.Formulas(12) = "pendingbooks='" & sss2 & "'"
    'cr1.Formulas(13) = "pendingbooks1='" & sss3 & "'"

    If MsgBox("Want to View ?", vbQuestion + vbYesNo) = vbYes Then
    
    
     cr1.WindowShowPrintSetupBtn = True
     cr1.WindowState = crptMaximized
     cr1.Action = 1
     d_ = 1
    
    Else
    
     cr1.Destination = crptToPrinter
     cr1.Action = 0
     d_ = 2
    End If

    
    
    
    cr1.Reset
    cr1.ReportFileName = rptPath & "/paidSlip.rpt"
    cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & invoice.I_NO.text & ""
    
    If invoice.cmbtransportname.text <> "" Then
       cr1.Formulas(0) = "transport_='" & invoice.cmbtransportname.text & "'"
    End If
    
    Dim add1, add2, add3 As String
    
    cr1.Formulas(1) = "shipto=''"
    If invoice.txtOrderNo.text <> "" Then
    If RS.State = 1 Then RS.Close
    RS.Open "select Shipto,Shipto_Add1,Shipto_Add2,shipto_dist,Shipto_States,bilty,ccattach,pin_ship from ordera where invoiceno=" & invoice.txtOrderNo.text & "", con
    If RS.EOF = False Then
       If RS!Shipto <> "" Then
          cr1.Formulas(1) = "shipto='SHIP TO :-'"
          cr1.Formulas(2) = "party='" & RS!Shipto & "'"
          
          add1 = RS!Shipto_Add1 & ""
          If Len(add1) > 0 Then
             If Len(RS!Shipto_Add2) > 0 Then
                add1 = add1 & ", " & RS!Shipto_Add2
             End If
          End If
          cr1.Formulas(3) = "add1_New='" & add1 & "'"
          If Len(RS!shipto_dist) > 0 Then
          cr1.Formulas(4) = "add2_New='" & RS!shipto_dist & "'"
          End If
          If Len(RS!Shipto_States) > 0 Then
          cr1.Formulas(5) = "add3_New='" & RS!Shipto_States & "'"
          End If
          
          If Not IsNull(RS!pin_ship) Then
             cr1.Formulas(10) = "pin_ship='" & "Pin : " & RS!pin_ship & "'"
          End If
          
          
       End If
          cr1.Formulas(6) = "paid_='" & RS!bilty & "'"
          cr1.Formulas(7) = "ccattach='" & RS!ccattach & "'"
          

    End If
    
    End If
    
    
     
    If d_ = 2 Then
       cr1.Destination = crptToPrinter
       cr1.Action = 0
    Else
     cr1.WindowShowPrintSetupBtn = True
     cr1.Action = 1
    
    End If
    '==============================
    
    
    
   
   
Else
    
    cr1.Reset
    cr1.ReportFileName = rptPath & "/invoice.rpt"
    cr1.Connect = "filedsn=chitradsn;uid=" & sql_user & ";pwd=" & sql_pass
    cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & invoice.I_NO.text & ""
    amt = (invoice.txtAmtwords.text)
    cr1.Formulas(2) = "RSS='" & amt & "'"
    cr1.Formulas(3) = "tin='" & tin & "'"
    cr1.WindowShowPrintSetupBtn = True
    cr1.WindowState = crptMaximized
    cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
    cr1.Formulas(8) = "unit_='" & cname_2 & "'"
    cr1.Formulas(9) = "inv='" & bill_ & "'"
    cr1.Formulas(10) = "gst='" & gst & "'"
    
    cr1.Formulas(11) = "scname_='" & invoice.txtSchool.text & "'"
    cr1.Formulas(12) = "sclbl='School'"
    cr1.Action = 1
    
 
    
    
'''s = ""
'''Dim oXApp As CRAXDRT.Application
'''Dim oXRpt As CRAXDRT.Report
'''Dim oXOpt As CRAXDRT.ExportOptions
'''
'''On Error GoTo ExportErr
'''
'''Set oXApp = CreateObject("CrystalRuntime.Application")
'''Set oXRpt = oXApp.OpenReport("" & rptPath & "/invoice.rpt")
'''oXRpt.RecordSelectionFormula = "{INVOICEA.invoiceno}=" & invoice.I_NO.Text & ""
'''With oXRpt
'''    .EnableParameterPrompting = False
'''    .MorePrintEngineErrorMessages = True
'''End With
'''
'''Set oXOpt = oXRpt.ExportOptions
'''
'''With oXOpt
'''    .DestinationType = crEDTDiskFile
'''    .DiskFileName = App.Path & "\abc.pdf"
'''    .FormatType = crEFTPortableDocFormat
'''End With
'''
'''oXRpt.export False  'throws missing or out-of-date dll error
'''ExportErr:
'''MsgBox err.Number & ", " & err.DESCRIPTION
'''

    
    
End If

PopUpValue6 = ""

ElseIf s1 = 2 Then

 cr1.Reset
 If frmBookIssueSp.Check1_trans.value = 0 Then
 cr1.ReportFileName = rptPath & "/invoice_sp.rpt"
 Else
   cr1.ReportFileName = rptPath & "/invoice_spTrans.rpt"
 End If
 
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & frmBookIssueSp.I_NO.text & ""
 'amt = toword(frmBookIssueSp.mna)
 amt = frmBookIssueSp.txtAmtwords.text
 
 cr1.Formulas(2) = "RSS='" & amt & "'"
 cr1.Formulas(3) = "tin='" & tin & "'"
 If RS.State = 1 Then RS.Close
 RS.Open "select Rep,Add1,Add2,Phone,pin,city,District,[state] from SalesRepQry where rep='" & frmBookIssueSp.cmbAgentName.text & "'", CON_blue
 If RS.EOF = False Then
    str_ = ""
    If RS!add1 <> "" Or Not IsNull(RS!add1) Then
       str_ = RS!add1
    End If
    If RS!add2 <> "" Or Not IsNull(RS!add2) Then
       str_ = str_ & ", " & RS!add2
    End If
    cr1.Formulas(4) = "add1_='" & str_1 & "'"
    
    str_ = ""
    If RS!phone <> "" And Not IsNull(RS!phone) Then
       str_ = RS!phone
    End If
    If RS!pin <> "" And Not IsNull(RS!pin) Then
       str_ = str_ & ", " & RS!pin
    End If
    cr1.Formulas(5) = "phone='" & str_ & "'"
    If UCase(RS!city) = UCase(RS!District) Then
       cr1.Formulas(6) = "city='" & UCase(RS!District) & "(" & RS!State & ")" & "'"
    Else
       cr1.Formulas(6) = "city='" & UCase(RS!city) & ", " & UCase(RS!District) & "(" & RS!State & ")" & "'"
    End If
    
 End If
 
 cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
 cr1.Formulas(8) = "unit_='" & cname_2 & "'"
 
 
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 If frmBookIssueSp.Check1_direct = 0 Then
   cr1.Action = 1
 Else
  cr1.Destination = crptToPrinter
  cr1.Action = 0


 End If
 
 
 '''''Paid Slip for Transport'''''''''''''''''
    '===================================
    If frmBookIssueSp.Check1_trans.value = 1 Then
    
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        
        cr1.Reset
        cr1.ReportFileName = rptPath & "/paidSlip_sp.rpt"
        cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
        cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & frmBookIssueSp.I_NO & ""
        
        
        If frmBookIssueSp.txtShip.text <> "" Then
           cr1.Formulas(0) = "party='" & frmBookIssueSp.txtShip.text & "'"
           cr1.Formulas(1) = "shipto='SHIP TO :-'"
           
             add1 = frmBookIssueSp.txtadd1 & ""
          If Len(add1) > 0 Then
             If Len(frmBookIssueSp.txtadd2) > 0 Then
                add1 = add1 & ", " & frmBookIssueSp.txtadd2
             End If
          End If
          cr1.Formulas(3) = "add1_New='" & add1 & "'"
          If Len(frmBookIssueSp.txtcity) > 0 Then
          cr1.Formulas(4) = "add2_New='" & frmBookIssueSp.txtcity & "'"
          End If
          If Len(frmBookIssueSp.txtstate) > 0 Then
          cr1.Formulas(5) = "add3_New='" & frmBookIssueSp.txtstate & "'"
          End If
           
           
        Else
           cr1.Formulas(0) = "party='" & frmBookIssueSp.cmbAgentName.text & "'"
           cr1.Formulas(1) = "shipto=''"
        End If
        If frmBookIssueSp.cmbtransportname.text <> "" Then
           cr1.Formulas(12) = "transport_='" & frmBookIssueSp.cmbtransportname.text & "'"
        End If
        
        If frmBookIssueSp.station.text <> "" Then
           If InStr(frmBookIssueSp.station.text, "BY") > 0 Then
              cr1.Formulas(13) = "station_='" & Trim(Mid(frmBookIssueSp.station.text, 1, InStr(frmBookIssueSp.station.text, "BY") - 1)) & "'"
           End If
        End If
        
        
        If RS.State = 1 Then RS.Close
        RS.Open "select ccattach from ordera where invoiceno=" & frmBookIssueSp.txtOrderNo.text & "", con
        If RS.EOF = False Then
        
        If Not IsNull(RS!ccattach) Then
          cr1.Formulas(15) = "ccattach='" & RS!ccattach & "'"
        End If
        
        End If
        
        If frmBookIssueSp.Check1_direct = 1 Then
           cr1.Destination = crptToPrinter
           cr1.Action = 0
        Else
           cr1.WindowShowPrintBtn = True
           cr1.Action = 1
        
        End If
    
    End If
    '==============================
 
 
 
 
 
ElseIf s1 = 3 Then

 cr1.Reset
 cr1.ReportFileName = rptPath & "/invoice_spret.rpt"
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & frmBookIssueSp_Ret.I_NO.text & ""
 amt = toword(frmBookIssueSp_Ret.mna)
 
 cr1.Formulas(2) = "RSS='" & amt & "'"
 cr1.Formulas(3) = "tin='" & tin & "'"
 If RS.State = 1 Then RS.Close
 RS.Open "select Rep,Add1,Add2,Phone,pin,city,District,[state] from SalesRepQry where rep='" & frmBookIssueSp_Ret.cmbAgentName.text & "'", CON_blue
 If RS.EOF = False Then
    str_ = ""
    If RS!add1 <> "" Or Not IsNull(RS!add1) Then
       str_ = RS!add1
    End If
    If RS!add2 <> "" Or Not IsNull(RS!add2) Then
       str_ = str_ & ", " & RS!add2
    End If
    cr1.Formulas(4) = "add1_='" & str_1 & "'"
    
    str_ = ""
    If RS!phone <> "" And Not IsNull(RS!phone) Then
       str_ = RS!phone
    End If
    If RS!pin <> "" And Not IsNull(RS!pin) Then
       str_ = str_ & ", " & RS!pin
    End If
    cr1.Formulas(5) = "phone='" & str_ & "'"
    If UCase(RS!city) = UCase(RS!District) Then
       cr1.Formulas(6) = "city='" & UCase(RS!District) & "(" & RS!State & ")" & "'"
    Else
       cr1.Formulas(6) = "city='" & UCase(RS!city) & ", " & UCase(RS!District) & "(" & RS!State & ")" & "'"
    End If
    
 End If
 
 cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
 cr1.Formulas(8) = "unit_='" & cname_2 & "'"
 
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Action = 1
 
 
ElseIf s1 = "4" Then
 '''Debit Note
 cr1.Reset
 
 If Debitnotefile.Check_header.value = 0 Then
    cr1.ReportFileName = rptPath & "/DebitNote.rpt"
 Else
    cr1.ReportFileName = rptPath & "/DebitNote_header.rpt"
 End If
 
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{DNFA.DNN}=" & Debitnotefile.TCNN.text & ""
 amt = (Debitnotefile.txtAmtwords.text)
 If Debitnotefile.Text3.text <> "" Then
   amt = toword(Debitnotefile.Text3.text)
 End If
 
 cr1.Formulas(1) = "RSS='" & amt & "'"
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Formulas(2) = "for_comp='" & cname_1 & "'"
 cr1.Formulas(3) = "unit_='" & cname_2 & "'"
 cr1.Action = 1
 
ElseIf s1 = "5" Then
 '''Credit Note
 cr1.Reset
 
 If Creditnotefile.Check1_withheader.value = 0 Then
    cr1.ReportFileName = rptPath & "/CreditNote.rpt"
 Else
    cr1.ReportFileName = rptPath & "/CreditNote_Header.rpt"
 End If
 
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{DNFA.CNN}=" & Creditnotefile.TCNN.text & ""
 amt = toword(Creditnotefile.Text3.text)
 cr1.Formulas(1) = "RSS='" & amt & "'"
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Formulas(2) = "for_comp='" & cname_1 & "'"
 cr1.Formulas(3) = "unit_='" & cname_2 & "'"
 cr1.Action = 1
 
ElseIf s1 = "6" Then
 '''Gledger report
 cr1.Reset
 cr1.ReportFileName = rptPath & "/GlAccount.rpt"
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{Winrpt.uid}=" & UId & ""
 cr1.Formulas(1) = "glacc='" & GLEDGERPRINT.COMBOGENLEDGER.text & "'"
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Action = 1
 
ElseIf s1 = "7" Then
 
 '''Subledger report
 
 cr1.Reset
 cr1.ReportFileName = rptPath & "/Al_SLAccount.rpt"
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{Winrpt.uid}=" & UId & ""
 cr1.Formulas(1) = "genledger='" & SLEDGERPRINT.COMBOGENLEDGER.text & "'"
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Formulas(2) = "fromdate='" & SLEDGERPRINT.date1.text & "'"
 cr1.Formulas(3) = "todate='" & SLEDGERPRINT.date2.text & "'"

 cr1.Action = 1
ElseIf s1 = "8" Then
 
 '''GL Trial report
 
 cr1.Reset
 cr1.ReportFileName = rptPath & "/GLTrialBalance.rpt"
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{Winrpt.uid}=" & UId & ""
 cr1.Formulas(1) = "asDate='" & Gentrial.date2.text & "'"
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized

 cr1.Action = 1

ElseIf s1 = "9" Then
 
 
 
    If subtrial.Check1_drcr.value = 1 Then
    
    cr1.Reset
    cr1.ReportFileName = rptPath & "/SLListClosing_DRCR.rpt"
    cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If subtrial.cboDRCr = "DR" Then
       cr1.ReplaceSelectionFormula "{winrpt.uid}= " & UId & " and {winrpt.dr}='Dr' and {winrpt.Closing}>0"
       cr1.Formulas(0) = "ledger= '" & subtrial.COMBOGENLEDGER.text & "'"
       cr1.Formulas(1) = "reporttitle= '" & "Debit Balance" & "'"
    ElseIf subtrial.cboDRCr = "CR" Then
       cr1.ReplaceSelectionFormula "({winrpt.uid}= " & UId & " and {winrpt.dr}='Cr' and {winrpt.Closing1}<0)"
       cr1.Formulas(0) = "ledger= '" & subtrial.COMBOGENLEDGER.text & "'"
       cr1.Formulas(1) = "reporttitle= '" & "Credit Balance" & "'"
    Else
       cr1.ReplaceSelectionFormula "{winrpt.uid}= " & UId & ""
       cr1.Formulas(0) = "ledger= '" & subtrial.COMBOGENLEDGER.text & "'"
       cr1.Formulas(1) = "reporttitle= '" & "Dr/Cr Balance" & "'"
    End If
    cr1.WindowShowExportBtn = True
    cr1.WindowShowPrintSetupBtn = True
    cr1.WindowShowPrintBtn = True
    cr1.WindowState = crptMaximized
    cr1.WindowShowRefreshBtn = True
    cr1.Action = 1
    
    Screen.MousePointer = vbDecimal
    
    Exit Sub
    
    End If
 
 
 '''SL Trial report
 
 cr1.Reset
 cr1.ReportFileName = rptPath & "/SLListClosing.rpt"
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.Formulas(1) = "reporttitle='" & subtrial.COMBOGENLEDGER.text & "'"
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized

 cr1.Action = 1

ElseIf s1 = "10" Then
 
 '''group wise sale
 cr1.Reset
 cr1.ReportFileName = rptPath & "/groupWiseSale.rpt"
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula PopUpValue1
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Action = 1
  
 PopUpValue1 = ""

ElseIf s1 = "11" Then
 
 '''group wise sale
 cr1.Reset
 cr1.ReportFileName = rptPath & "/groupWiseSaleret.rpt"
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula PopUpValue1
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Action = 1

ElseIf s1 = "12" Then

If Critnote.Check1_direct.value = 0 Then

 cr1.Reset
 
 If Critnote.Check_header.value = 0 Then
    cr1.ReportFileName = rptPath & "/CreditNotItem.rpt"
 Else
    cr1.ReportFileName = rptPath & "/CreditNotItem_Head.rpt"
 End If
 
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & Critnote.I_NO.text & ""
 amt = (Critnote.txtAmtwords.text)
 cr1.Formulas(2) = "RSS='" & amt & "'"
 cr1.Formulas(3) = "tin='" & tin & "'"
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
 cr1.Formulas(8) = "unit_='" & cname_2 & "'"
 cr1.Action = 1

Else

 DoEvents
 DoEvents
 DoEvents

 cr1.Reset
 
 If Critnote.Check_header.value = 0 Then
    cr1.ReportFileName = rptPath & "/CreditNotItem.rpt"
 Else
    cr1.ReportFileName = rptPath & "/CreditNotItem_Head.rpt"
 End If
 
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & Critnote.I_NO.text & ""
 amt = (Critnote.txtAmtwords.text)
 cr1.Formulas(2) = "RSS='" & amt & "'"
 cr1.Formulas(3) = "tin='" & tin & "'"
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
 cr1.Formulas(8) = "unit_='" & cname_2 & "'"
 cr1.Destination = crptToPrinter
 cr1.Action = 0
 
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 
 cr1.Reset
 
 If Critnote.Check_header.value = 0 Then
    cr1.ReportFileName = rptPath & "/CreditNotItem.rpt"
 Else
    cr1.ReportFileName = rptPath & "/CreditNotItem_Head.rpt"
 End If
 
 cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
 cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & Critnote.I_NO.text & ""
 amt = (Critnote.txtAmtwords.text)
 cr1.Formulas(2) = "RSS='" & amt & "'"
 cr1.Formulas(3) = "tin='" & tin & "'"
 cr1.WindowShowPrintSetupBtn = True
 cr1.WindowState = crptMaximized
 cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
 cr1.Formulas(8) = "unit_='" & cname_2 & "'"
 cr1.Destination = crptToPrinter
 cr1.Action = 0
 






End If



ElseIf s1 = "101" Then

 
    DoEvents
    DoEvents
 
    
    cr1.Reset
    cr1.ReportFileName = rptPath & "/invoice.rpt"
    cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & invoice.I_NO.text & ""
    amt = (invoice.txtAmtwords.text)
    cr1.Formulas(2) = "RSS='" & amt & "'"
    cr1.Formulas(3) = "tin='" & tin & "'"
    cr1.WindowShowPrintSetupBtn = True
    cr1.WindowState = crptMaximized
    cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
    cr1.Formulas(8) = "unit_='" & cname_2 & "'"
    cr1.Formulas(9) = "inv='" & bill_ & "'"
    cr1.Formulas(10) = "gst='" & gst & "'"
    cr1.Formulas(11) = "scname_=''"
    cr1.Formulas(12) = "sclbl=''"
    cr1.Destination = crptToPrinter
    cr1.Action = 0

    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents

    cr1.Reset
    cr1.ReportFileName = rptPath & "/invoice.rpt"
    cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & invoice.I_NO.text & ""
    amt = (invoice.txtAmtwords.text)
    cr1.Formulas(2) = "RSS='" & amt & "'"
    cr1.Formulas(3) = "tin='" & tin & "'"
    cr1.WindowShowPrintSetupBtn = True
    cr1.WindowState = crptMaximized
    cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
    cr1.Formulas(8) = "unit_='" & cname_2 & "'"
    cr1.Formulas(9) = "inv='" & bill_ & "'"
    cr1.Formulas(10) = "gst='" & gst & "'"
    cr1.Formulas(11) = "scname_='" & invoice.txtSchool.text & "'"
    cr1.Formulas(12) = "sclbl='School'"
    
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    
    cr1.Destination = crptToPrinter
    cr1.Action = 0
    


    
    
ElseIf s1 = "200" Then
     
    bill_ = billformat & "" & Format(Trim(Str(countersale.I_NO)), "00000")

    cr1.Reset
    cr1.ReportFileName = rptPath & "/invoice_cash.rpt"
    cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr1.ReplaceSelectionFormula "{INVOICEA.invoiceno}=" & countersale.I_NO.text & ""
    'amt = (countersale.txtAmtwords.Text)
    'cr1.Formulas(2) = "RSS='" & amt & "'"
    cr1.Formulas(3) = "tin='" & tin & "'"
    cr1.WindowShowPrintSetupBtn = True
    cr1.WindowState = crptMaximized
    cr1.Formulas(7) = "for_comp='" & cname_1 & "'"
    cr1.Formulas(8) = "unit_='" & cname_2 & "'"
    cr1.Formulas(9) = "inv='" & bill_ & "'"
    cr1.Formulas(10) = "gst='" & gst & "'"
    cr1.Formulas(11) = "scname_='" & countersale.txtSchool.text & "'"
    cr1.Formulas(12) = "sclbl='School'"
    
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    
    cr1.WindowShowPrintSetupBtn = True
    cr1.WindowState = crptMaximized
    cr1.Action = 1
    
    'cr1.Destination = crptToPrinter
    
    

End If




Screen.MousePointer = vbDefault


End Sub
Private Sub Command2_Click()
Dim V As String

         

viewinvoice.Left = 0
viewinvoice.top = 10
viewinvoice.Show
Unload Me


End Sub
Private Sub Command3_Click()

Dim address1, address2, address3, address4 As String
Dim through, transport, bundle As String


If s1 = 1 Then


    
    If invoice.lblMail.Caption = "" Then
       MsgBox "Mail Can'nt Be Send..,Because Party Mail Id is not Found "
       Exit Sub
    End If
    
    'For Invoice
    Dim rep_email As String

         
    If RS.State = 1 Then RS.Close
    RS.Open "select party,ADDRESS1,ADDRESS2,district,states,THROUGH,THROUGH1,BUNDLES,STATION from invoiceaQry where invoiceno=" & invoice.I_NO & "", con
    If RS.EOF = False Then
       address1 = RS!party
       address2 = RS!address1
       address3 = RS!address2
       address4 = RS!District + ",(" + RS!states + ")"
       through = RS!through + " " + RS!through1
       bundle = RS!bundles & ""
       transport = RS!station & ""
       
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "select Email FROM Rep where rep='" & invoice.cmbAgentName & "'", CON_blue
    If RS.EOF = False Then
      If (RS!email <> "" Or Not IsNull(RS!email)) Then
         rep_email = RS!email
      Else
         rep_email = "-"
      End If
    End If
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "select * from MailDetails where (Bill=" & invoice.I_NO & " and BillType='invoice')", con
    If RS.EOF = True Then
       con.Execute "insert into MailDetails(Bill,BillType,Mail,MailSended,OrderDis,address1,address2,address3,address4,through,Transport,Bundle,RepEmail)" & _
       " values(" & invoice.I_NO & ",'invoice','" & invoice.lblMail & "','y','n','" & address1 & "','" & address2 & "','" & address3 & "','" & address4 & "','" & through & "','" & transport & "','" & bundle & "','" & rep_email & "')"
    Else
       con.Execute "update MailDetails set MailSended='y',OrderDis='n',Mail='" & invoice.lblMail & "',RepEmail='" & rep_email & "' where (Bill=" & invoice.I_NO & " and BillType='invoice')"
    End If
    
    
ElseIf s1 = 2 Then


    If RS.State = 1 Then RS.Close
    RS.Open "SELECT INVOICENO,MARKA,BUNDLES,THROUGH,STATION,THROUGH1,AgentName,Add1,Add2,district,City,States,mail FROM INVOICEA_sp where invoiceno=" & frmBookIssueSp.I_NO.text & "", con
    If RS.EOF = False Then
       address1 = RS!agentname
       address2 = RS!add1
       address3 = RS!add2
       address4 = RS!District + ",(" + RS!states + ")"
       through = RS!through + " " + RS!through1
       bundle = RS!bundles & ""
       transport = RS!station & ""
       mail = RS!mail & ""
       
    End If
    
    If mail = "" Then
       MsgBox "Mail Can'nt Be Send..,Because Representative Mail Id is not Found "
       Exit Sub
    End If

    
    If RS.State = 1 Then RS.Close
    RS.Open "select * from MailDetails where (Bill=" & frmBookIssueSp.I_NO & " and BillType='invoice_sp')", con
    If RS.EOF = True Then
       con.Execute "insert into MailDetails(Bill,BillType,Mail,MailSended,OrderDis,address1,address2,address3,address4,through,Transport,Bundle)" & _
       " values(" & frmBookIssueSp.I_NO.text & ",'invoice_sp','" & mail & "','y','n','" & address1 & "','" & address2 & "','" & address3 & "','" & address4 & "','" & through & "','" & transport & "','" & bundle & "')"
    Else
       con.Execute "update MailDetails set MailSended='y',OrderDis='n',Mail='" & mail & "' where (Bill=" & frmBookIssueSp.I_NO & " and BillType='invoice_sp')"
    End If

    
    
End If


MsgBox "Mail has been send...", vbInformation

End Sub

Private Sub Command4_Click()


'Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
'myURL = "http://crm.blueprinteducation.co.in/shootwhatsapp.php?invId=INEG10024&mobile=8394888882&dear=Nitin "
'winHttpReq.Open "POST", myURL, False
'winHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'winHttpReq.Send
'SendSMS = winHttpReq.responseText
'


If s1 = 0 Then
   MsgBox "Error...", vbCritical
   Exit Sub
End If

Dim rep_email, head_mail, mob As String
Dim M_Mailid As String

Dim rs_ As New ADODB.Recordset
Dim rs_mn As New ADODB.Recordset

Dim randomId As String
Dim s10 As String
mob = ""
randomId = ""
M_Mailid = ""

Dim inv_No As String

s1 = 1

'INVOICEA_sp

If rs1.State = 1 Then rs1.Close
rs1.Open "select invoiceno  from INVOICEA_sp where RandomId is null order by invoiceno", con

While rs1.EOF = False



If RS.State = 1 Then RS.Close
RS.Open "select RandomNo from INVOICEA_sp where invoiceno=" & rs1!invoiceNo & "", con
If RS.EOF = False Then
   
   inv_No = ""
   
   inv_No = RS!RandomNo

    If Len(inv_No) >= 5 Then
       s10 = Mid(value_, 1, 2)
    Else
       s10 = Mid(value_, 1, 3)
    End If
    
    s10 = "IN" & s10 & inv_No
        
    con.Execute "update INVOICEA_sp set Randomid= '" & s10 & "' where invoiceno=" & rs1!invoiceNo & ""
 End If

    
rs1.MoveNext
    
Wend




End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Dim V As String


'login.DSN
   
DSNNew
    
    
    
 cmdprinter.Enabled = True
 If (Val(s1) > 3 Or Val(s1) = 5) Then
 cmdprinter.Enabled = False
 End If
 
 If (Val(s1) = 5) Then
 Command2.Enabled = False
 End If
 
 If (Val(s1) = 12 Or Val(s1) = 7) Then
    cmdprinter.Enabled = True
 End If
 
 
 If printButton = "1" Then
   Command2.Enabled = False
Else
   Command2.Enabled = True
End If

 
 

End Sub


