VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDisPatchDaily 
   Caption         =   "Dispatch Report"
   ClientHeight    =   2904
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5724
   Icon            =   "frmDisPatchDaily.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2904
   ScaleWidth      =   5724
   Begin VB.CommandButton cmdMail 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Send Mail"
      Height          =   744
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1728
      Width           =   1356
   End
   Begin Crystal.CrystalReport cr 
      Left            =   5184
      Top             =   540
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   744
      Left            =   3744
      Picture         =   "frmDisPatchDaily.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1728
      Width           =   1356
   End
   Begin VB.CommandButton cmdPrint_7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   744
      Left            =   936
      Picture         =   "frmDisPatchDaily.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1728
      Width           =   1356
   End
   Begin MSComCtl2.DTPicker txtDT 
      Height          =   396
      Left            =   1980
      TabIndex        =   2
      Top             =   792
      Width           =   1536
      _ExtentX        =   2709
      _ExtentY        =   699
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   118554625
      CurrentDate     =   42409
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1188
      TabIndex        =   3
      Top             =   828
      Width           =   876
   End
End
Attribute VB_Name = "frmDisPatchDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdMail_Click()

'=====================================================================================

Dim rs_ As New ADODB.Recordset
Dim rs1_ As New ADODB.Recordset

Dim mailId As String


If Len(Day(txtDT.value)) = 1 Then
   mailId = "9" & Trim(Day(txtDT.value))
Else
   mailId = "" & Trim(Day(txtDT.value))
End If

'mailId = Trim(Day(txtDT.value))

If Len(Month(txtDT.value)) = 1 Then
    mailId = mailId & "0" & Trim(Month(txtDT.value))
Else
    mailId = mailId & "" & Trim(Month(txtDT.value))
End If

mailId = mailId & "0" & Right(Trim(Year(txtDT.value)), 2)


Dim HeadEmail_2 As String
Dim totQty As Long
totQty = 0

HeadEmail_2 = ""

If rs1.State = 1 Then rs1.close
rs1.Open "select invoiceno,party,agentname,CreatedDt from invoiceaQry " & _
" where convert(date,CreatedDt,103)=convert(datetime,'" & txtDT.value & "',103) " & _
"and party not in('FLIPKART INTERNET (P) LTD.','AMAZON SELLER SERVICES PVT.LTD.','ONLINE ORDER C/O AVENUE INDIA PVT.LTD.') " & _
"order by invoiceno", con
While rs1.EOF = False

   If rs_.State = 1 Then rs_.close
   rs_.Open "select headname,HeadEmail,phone,HeadEmail from Rep where Rep='" & rs1!agentname & "' AND LEN(HeadEmail)>0", CON_blue
   If rs_.EOF = False Then
   
      HeadEmail_2 = rs_!HeadEmail & ""
      
      If RS.State = 1 Then RS.close
      RS.Open "select headname,HeadEmail,phone,headname from Rep where rep='" & rs_!headName & "'", CON_blue
      If RS.EOF = False Then
         hmobile = Mid(RS!phone, 1, 10)
      End If
      
     If rs1_.State = 1 Then rs1_.close
     rs1_.Open "SELECT sum(QUANTITY) as qty FROM INVOICEB  where invoiceno=" & rs1!invoiceNo & "", con
     If Not IsNull(rs_(0)) Then
        totQty = totQty + rs1_(0)
     End If

      
      con.Execute "update invoicea set HeadName2='" & HeadEmail_2 & "',headnames='" & rs_!headName & "',dispatchmailid='" & mailId & "',HeadMobile='" & hmobile & "',tqty=" & totQty & " where invoiceno=" & rs1!invoiceNo & ""
   End If
 
   rs1.MoveNext
Wend



''=====================================================================================





If MsgBox("Want to Send Mail .. ?", vbQuestion + vbYesNo) = vbYes Then


Set rs_ = New ADODB.Recordset
'Dim mailId As String
Dim headMail As String

k1 = 1




Dim manager As String
Dim headMail_2 As String
Dim rs2 As New ADODB.Recordset
Dim cdt
Dim pname As String

If rs1.State = 1 Then rs1.close
rs1.Open "select headnames,HeadMobile from invoiceaQry " & _
" where dispatchMailId='" & mailId & "'" & _
"group by headnames,HeadMobile", con

If rs1.EOF = True Then
   MsgBox "Record Not found ....", vbInformation
   Exit Sub
End If

kk1 = 1


While rs1.EOF = False

   headMail_2 = ""
   pname = ""
   
   If rs_.State = 1 Then rs_.close
   rs_.Open "select headname,HeadEmail,ManagerMail,HeadEmail_2 from Rep where headname='" & rs1!headnames & "' AND LEN(HeadEmail)>0", CON_blue
   If rs_.EOF = False Then
      
      manager = rs_!ManagerMail & ""
      headMail = rs_!HeadEmail
      
      If Len(k1) = 1 Then
         bill_ = mailId & "0" & k1
      Else
         bill_ = mailId & "" & k1
      End If
      
      headMail_2 = ""
      
      If rs2.State = 1 Then rs2.close
      rs2.Open "select HeadEmail from Rep where HeadName_2='" & rs_!headName & "'", CON_blue
      If rs2.EOF = False Then
          headMail_2 = rs2!HeadEmail & ""
      End If
      
      If kk1 = 1 Then
         pname = "nitinrastogi"
      Else
         pname = ""
         totQty = 0
      End If
      
      kk1 = kk1 + 1
            
      If RS.State = 1 Then RS.close
      RS.Open "select * from MailDetails where (Bill=" & bill_ & " and BillType='dispatch-mail' and mail='" & headMail & "')", con
      If RS.EOF = True Then
         
         con.Execute "insert into MailDetails(Bill,BillType,Mail,MailSended,HeadEmail,Manager,HeadEmail_2,through,CreatedDt,partyname,profile_)" & _
         " values(" & bill_ & ",'dispatch-mail','" & headMail & "','n','" & manager & "','" & rs_!headName & "','" & headMail_2 & "','" & rs1!HeadMobile & "','" & Format(txtDT.value, "MM/dd/yyyy") & "','" & pname & "'," & totQty & ")"
    
    
     End If
    
     k1 = k1 + 1
     
   End If
 
   rs1.MoveNext
Wend


MsgBox "Mail Sended..."

End If



End Sub
Private Sub cmdPrint_7_Click()

Dim rs_ As New ADODB.Recordset
Dim mailId As String

Dim totQty As Long

totQty = 0


If Len(Day(txtDT.value)) = 1 Then
   mailId = "0" & Trim(Day(txtDT.value))
Else
   mailId = "" & Trim(Day(txtDT.value))
End If

If Len(Month(txtDT.value)) = 1 Then
    mailId = mailId & "0" & Trim(Month(txtDT.value))
Else
    mailId = mailId & "" & Trim(Month(txtDT.value))
End If

mailId = mailId & Trim(Year(txtDT.value))




If rs1.State = 1 Then rs1.close
rs1.Open "select invoiceno,party,agentname,CreatedDt from invoiceaQry " & _
" where convert(date,CreatedDt,103)=convert(datetime,'" & txtDT.value & "',103) " & _
"and party not in('FLIPKART INTERNET (P) LTD.','AMAZON SELLER SERVICES PVT.LTD.','ONLINE ORDER C/O AVENUE INDIA PVT.LTD.') " & _
"order by invoiceno", con
While rs1.EOF = False

   If rs_.State = 1 Then rs_.close
   rs_.Open "select headname,HeadEmail,phone,headname from Rep where Rep='" & rs1!agentname & "' AND LEN(HeadEmail)>0", CON_blue
   If rs_.EOF = False Then
      
      If RS.State = 1 Then RS.close
      RS.Open "select headname,HeadEmail,phone,headname from Rep where rep='" & rs_!headName & "'", CON_blue
      If RS.EOF = False Then
      hmobile = Mid(RS!phone, 1, 10)
      End If
      
      con.Execute "update invoicea set headnames='" & rs_!headName & "',dispatchmailid='" & mailId & "',HeadMobile='" & hmobile & "' where invoiceno=" & rs1!invoiceNo & ""
   
   End If
   
   If rs_.State = 1 Then rs_.close
   rs_.Open "SELECT sum(QUANTITY) as qty FROM INVOICEB  where invoiceno=" & rs1!invoiceNo & "", con
   If Not IsNull(rs_(0)) Then
      totQty = totQty + rs_(0)
   End If

  
   rs1.MoveNext
Wend



''=====================================================================================

Dim inv1
k1 = 0

If rs1.State = 1 Then rs1.close
rs1.Open "select invoiceno,party from invoiceaQry " & _
" where convert(date,CreatedDt,103)=convert(datetime,'" & txtDT.value & "',103) " & _
"and party not in('FLIPKART INTERNET (P) LTD.','AMAZON SELLER SERVICES PVT.LTD.','ONLINE ORDER C/O AVENUE INDIA PVT.LTD.') " & _
"order by invoiceno", con
While rs1.EOF = False
   
   If k1 = 0 Then
      inv1 = "{invoiceaQry.invoiceNo} = " & rs1!invoiceNo & ""
      k1 = k1 + 1
   Else
      inv1 = inv1 & " or {invoiceaQry.invoiceNo} = " & rs1!invoiceNo & ""
   End If
 
   rs1.MoveNext
   
Wend

If IsEmpty(inv1) Then
   MsgBox "Record Not found ....", vbInformation
   Exit Sub
End If



DSNNew


MsgBox "Want to Print .. ?", vbQuestion + vbYesNo

cr.Reset
cr.ReportFileName = rptPath & "\dispatchreport.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.ReplaceSelectionFormula "{invoiceaQry.party}<>'FLIPKART INTERNET (P) LTD.' and " & inv1

cr.Formulas(0) = "dt='" & txtDT.value & "'"
cr.Formulas(1) = "tqty=" & totQty & ""


cr.WindowShowPrintSetupBtn = True
cr.WindowShowPrintBtn = True
cr.WindowState = crptMaximized
cr.Action = 1

End Sub
Private Sub Form_Load()
  txtDT.value = Format(Date, "dd/MM/yyyy")
  
  Me.Width = 6012
  Me.Height = 3348
End Sub
