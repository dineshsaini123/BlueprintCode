VERSION 5.00
Begin VB.Form frmSendMail 
   Caption         =   "Mail Option..."
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5736
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5736
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lblmail3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   5460
   End
   Begin VB.TextBox lblmail2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2340
      Width           =   5460
   End
   Begin VB.TextBox lblmail1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   5460
   End
   Begin VB.TextBox lblmail 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Width           =   5460
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "&Send Mail"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   5475
   End
   Begin VB.ComboBox cbomail 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Select ........"
      Top             =   900
      Width           =   5520
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   120
      TabIndex        =   8
      Top             =   15
      Width           =   5565
      Begin VB.ComboBox cboProfile 
         Height          =   315
         Left            =   4275
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.OptionButton Option1_toallparty 
         Caption         =   "To All Party"
         Height          =   435
         Left            =   2970
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.OptionButton Option2_party 
         Caption         =   "To Party"
         Height          =   435
         Left            =   1905
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton Option1_rep 
         Caption         =   "To Representative"
         Height          =   315
         Left            =   135
         TabIndex        =   0
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Profile"
         Height          =   255
         Index           =   9
         Left            =   4275
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbomail_Click()

If Option1_rep.value = True Then

 If RS.State = 1 Then RS.close
 RS.Open "select email from SalesRepQry where rep='" & cbomail & "'", CON_blue
 If RS.EOF = False Then
    lblMail.Text = RS!email
 End If

Else

 If RS.State = 1 Then RS.close
 RS.Open "select email from sledger where subledger='" & cbomail & "'", con
 If RS.EOF = False Then
    lblMail.Text = RS!email
 End If


End If

End Sub
Sub addData(rptid As String, rptid1 As String)
 con.Execute "insert into tempLedger7(dates,Billtype,Bill,Des,Dr,Cr,Balance,drcr,Party,OpBalance,setupid,fyear,UserId,rptid,rptype,rptName,states,RepName,Party1,AspectedAmt)" & _
 "SELECT dates,Billtype,Bill,Des,Dr,Cr,Balance,drcr,Party,OpBalance,setupid,fyear,UserId,'" & rptid1 & "',rptype,rptName,states,RepName,Party1,AspectedAmt from tempLedger7 where rptid='" & rptid & "'"
End Sub

Private Sub cboProfile_Click()
cmdSendMail.Enabled = True
End Sub
Private Sub cmdSendMail_Click()

Dim rptid As String
rptid = ""

'-1
Dim add1, add4 As String

If rs1.State = 1 Then rs1.close
rs1.Open "select party,address3 from SLEDGER where subledger='" & PopUpValue6 & "'", con
If rs1.EOF = False Then
   add1 = rs1!party & ""
   add4 = rs1!address3 & ""
   
End If



''''ADd Data In tempLedger7================================================
If RS.State = 1 Then RS.close
RS.Open "select max(bill) from MailDetails", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   
   If IsNull(RS(0)) Then
      popupvalue5 = 9999 + 1
   Else
      popupvalue5 = RS(0) + 1
   End If
   
   con.Execute ("exec PartyStateMentPartucullar '" & PopUpValue6 & "','" & UId & "'")
   con.Execute "update tempLedger7 set rptid=" & popupvalue5 & " where party='" & PopUpValue6 & "'"
 End If
PopUpValue6 = ""
'==========================================================================



If lblMail <> "" Then
    
    rptid = popupvalue5
    
    If RS.State = 1 Then RS.close
    RS.Open "select Bill from MailDetails where bill='" & popupvalue5 & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
      con.Execute "insert into MailDetails(Bill,Mail,BillType,rptName,address1,address4,partyname) values('" & popupvalue5 & "','" & lblMail & "','ledger','" & popupvalue4 & "','" & add1 & "','" & add4 & "','" & frmBillList.cboParty & "')"
      cmdSendMail.Enabled = False
    Else
      con.Execute "update MailDetails set MailSended='n',rptName='" & popupvalue4 & "' where Bill='" & popupvalue5 & "'"
      cmdSendMail.Enabled = False
    End If
    
End If

'-2
If lblmail1 <> "" Then
    popupvalue5 = popupvalue5 + 1
    addData rptid, popupvalue5
    
    If RS.State = 1 Then RS.close
    RS.Open "select Bill from MailDetails where bill='" & popupvalue5 & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
      con.Execute "insert into MailDetails(Bill,Mail,BillType,rptName,address1,address4) values('" & popupvalue5 & "','" & lblmail1 & "','ledger','" & popupvalue4 & "','" & add1 & "','" & add4 & "')"
      cmdSendMail.Enabled = False
    End If
End If

'-3
If lblmail2 <> "" Then
    popupvalue5 = popupvalue5 + 1
    addData rptid, popupvalue5
    
    If RS.State = 1 Then RS.close
    RS.Open "select Bill from MailDetails where bill='" & popupvalue5 & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
      con.Execute "insert into MailDetails(Bill,Mail,BillType,rptName,address1,address4) values('" & popupvalue5 & "','" & lblmail2 & "','ledger','" & popupvalue4 & "','" & add1 & "','" & add4 & "')"
      
      cmdSendMail.Enabled = False
    End If
End If

'-4
If lblmail3 <> "" Then
    popupvalue5 = popupvalue5 + 1
    addData rptid, popupvalue5
    
    If RS.State = 1 Then RS.close
    RS.Open "select Bill from MailDetails where bill='" & popupvalue5 & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
      con.Execute "insert into MailDetails(Bill,Mail,BillType,rptName,address1,address4) values('" & popupvalue5 & "','" & lblmail3 & "','ledger','" & popupvalue4 & "','" & add1 & "','" & add4 & "')"
      cmdSendMail.Enabled = False
    End If
End If



popupvalue5 = ""


Dim jj As Long

If Option1_toallparty.value = True Then
   con.Execute ("exec PartyStateMentnew " & UId & "")
   jj = 91111
   con.Execute "delete from MailDetails where (profile_='" & cboProfile & "' and rptname='PartyLedgerBillwise.rpt')"
   If RS.State = 1 Then RS.close
   RS.Open "SELECT DISTINCT SLEDGER.email,SLEDGER.party AS ADDRESS1,'ledger',SLEDGER.ADDRESS1 AS ADDRESS2,SLEDGER.ADDRESS2 AS ADDRESS3,SLEDGER.ADDRESS3 AS ADDRESS4,tempLedger_allst.rptid,tempLedger_allst.Party FROM tempLedger_allst INNER JOIN SLEDGER ON tempLedger_allst.Party = SLEDGER.SUBLEDGER where SLEDGER.profile_='" & cboProfile.Text & "'", con
   While RS.EOF = False
      
      If Len(RS(0)) > 5 Then
         'con.Execute "insert into MailDetails(bill,Mail,Address1,BillType,Address2,Address3,Address4,rptname,profile_) values('" & RS(6) & "','" & RS(0) & "','" & RS(1) & "','" & RS(2) & "','" & RS(3) & "','" & RS(4) & "','" & RS(5) & "','PartyLedgerBillwise.rpt','" & cboProfile.Text & "') "
         con.Execute "update tempLedger_allst set rptid=" & jj & " where party='" & RS(7) & "'"
         con.Execute "insert into MailDetails(bill,Mail,Address1,BillType,Address2,Address3,Address4,rptname,profile_,MailSended) values('" & jj & "','" & RS(0) & "','" & RS(1) & "','" & RS(2) & "','" & RS(3) & "','" & RS(4) & "','" & RS(5) & "','PartyLedgerBillwise.rpt','" & cboProfile.Text & "','Bulk Mail...') "
      jj = jj + 1
      End If
      
      
      RS.MoveNext
   Wend
   
   con.Execute "delete from tempLedger_allst where rptid=1"
   cmdSendMail.Enabled = False
End If




MsgBox ("Mail Sended (Process Continue..)")


End Sub

Private Sub Form_Load()
add_

cbomail.Text = PopUpValue6
lblMail.Text = PopUpValue3

PopUpValue3 = ""
PopUpValue6 = ""

End Sub
Sub add_()
  
cbomail.Clear

cbomail.Enabled = True
lblMail.Enabled = True
lblmail1.Enabled = True
lblmail2.Enabled = True

Label20(9).Visible = False
cboProfile.Visible = False

  
If Option1_rep.value = True Then

  If RS.State = 1 Then RS.close
  RS.Open "select Rep as Representative from SalesRepQry where len(Email)>0 order by Rep", CON_blue
  
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cbomail.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If

ElseIf Option2_Party.value = True Then

  'If RS.State = 1 Then RS.close
  'RS.Open "select subledger from sledger  order by subledger", con
  Set RS = New ADODB.Recordset
  Set RS = con.Execute("exec searchList 'party'")
  
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cbomail.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If

Else

cbomail.Enabled = False
lblMail.Enabled = False
lblmail1.Enabled = False
lblmail2.Enabled = False


'======================================
If RS.State = 1 Then RS.close
'RS.Open "select distinct Profile_ from SLEDGER order by Profile_", con
Set RS = New ADODB.Recordset
Set RS = con.Execute("exec searchList 'profile'")

If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cboProfile.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If


Label20(9).Visible = True
cboProfile.Visible = True

    
End If
    
    
End Sub

Private Sub Option1_rep_Click()
add_
End Sub

Private Sub Option1_toallparty_Click()
add_
End Sub

Private Sub Option2_party_Click()
add_
End Sub
