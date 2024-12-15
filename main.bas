Attribute VB_Name = "main"
Public con As ADODB.Connection               'Data.mdb
Public CCON As ADODB.Connection              'Config
Public CONDSN As ADODB.Connection
Public CON_noida As ADODB.Connection
Public CON_blue As ADODB.Connection
Public con_LAST As ADODB.Connection
Public con_LAST2 As ADODB.Connection
Public con_doc1 As New ADODB.Connection
'Public con_LAST1 As ADODB.Connection
Public database_last As String
Public database_last_ As String
Public coninfo As ADODB.Connection
Public LocaldatabasePath As String
Public rptPath As String
Public bookOp As String

Public party_name As String

Public v_vtype As String
Public v_vdate As String
Public v_vnumber As String
Public v_Remarks As String


Public sql_pass As String
Public sql_user As String

Public current_dbase As String
Public last_dbase  As String
Public next_dbase As String
Public ch_din  As String
Public donnation_visible  As String
Public printButton  As String
Public financialyear_Fdate As Date
Public financialyear_Tdate As Date

Public financialyear_Fdate_SaleRet As Date
Public financialyear_Tdate_SaleRet As Date

Public vtypeNew As String
Public vtype1_ As String
Public vdate_ As Date
Public vno_  As String
Public actionType_  As String

Public AuditTrail  As String
Public ledger_ As String
Public inv_ledger  As String
Public pname_ As String

Public Sys_user_ As String
Public databaseNew As String
Public serverNameNew As String
Public billformat  As String
Public itmeCode  As String
Public gst As String
Public strledger As String
Public firm As String
Public col_search As String
Public com_name, com_user As String
Public donnation_ As String
Public frmNo As String
Public mnuMenu_ As String
Public debitForAgn As String
Public debitForAgnNew As String
Public SessionLastDate As Date



Public serverName_ As String
Public session_next As String
Public fromDate_setup, toDate_setup As String
'Globle varible for session date
Public server_ As String
Public current_next As String

Public rpt_type As String
Public cname_1 As String
Public cname_2 As String
Public cname_add1 As String
Public packing_ As String
Public UserName As String
Public from_date, to_date As Date

Public PopUpValue1 As String
Public PopUpValue2 As String
Public PopUpValue3 As String

Public popupvalue4 As String
Public popupvalue5 As String
Public PopUpValue6 As String
Public PopUpValue7 As String

Public HeadTbl As String
Public cp As String
Public module_ As String
Public searchby As String
Public sqlQry As String
Public orderby As String
Public setupid As Byte
Public inviceNo As String

Public session As String
Public session_last1 As String
Public session_last2 As String

Public reportname As String
Public directory As String
Public repocon As ADODB.Connection
Public repors As ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public RS As New ADODB.Recordset
Public firm_Address As String
Public searchType As String
Public UId As Integer
Public ss1 As Integer
Public s2 As Integer
Public s1 As Integer
Public searchForm As String
Public bookNo As String
Public stringyear As String
Public blnviewallcomp As Boolean
Public strrptpath As String
Public strinvrpt As String
Public arycname() As String
Public constr As String
Public tblNo As String
Public export_flage As Boolean
Public export_currency As String
Public headData As String
Public Paper_Master  As String
Public d10 As Integer
Public rtype_ As Integer
Public vtypes, vdates, vnumbers As String

'----------------New  Variable -----------

Public IssueBook As String
Public change_Pass As String
Public Rec1 As String
Public hindi As String
Public english As String
Public ream_tot, sheet_tot

Public leftAlign, leftAlign_cash As Integer
Public leftAlign_crnot_I As Integer
Public Const exportsale = "Export Sales"
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
     (ByVal lpBuffer As String, nSize As Long) As Long
Public Sub sendkeys(text As Variant, Optional wait As Boolean = False)

Dim wshshell As Object
Set wshshell = CreateObject("wscript.shell")
wshshell.sendkeys CStr(text), wait
Set wshshell = Nothing

End Sub
'----------------End  Code ---------------
Public Function AuditTrail_Log(vt As String, ActionType As String, ReasionForEdit As String, vdate As Date, vno As String, vtype As String) As Boolean

Dim rss As New ADODB.Recordset
Set rss = New ADODB.Recordset

If vt = "V" Then
   
If (ActionType = "Insert") Then
     rss.Open "select VoucherID,VoucherType,VoucherDate,VoucherNumber,Amount,Particular from  VOUCHERS_Main where vouchertype='" + Trim(vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And VoucherID = " + vno, con
   Else
     rss.Open "select VoucherID,VoucherType,VoucherDate,VoucherNumber,Amount,Particular from  VOUCHERS_Main where vouchertype='" + Trim(vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate & "',103) And vouchernumber = " + vno, con
   End If
   
ElseIf vt = "I" Then
   
   rss.Open "select INVOICENO as VoucherID,'Sale Invoice' as VoucherType,INVOICEDATE as VoucherDate,0 as VoucherNumber,netamount as Amount,SUBLEDGER as Particular from  invoicea where invoiceno = " + vno, con

ElseIf vt = "CI" Then
   
   rss.Open "select INVOICENO as VoucherID,'Credit Note (Item)' as VoucherType,INVOICEDATE as VoucherDate,0 as VoucherNumber,netamount as Amount,SUBLEDGER as Particular from  credita where invoiceno = " + vno, con

ElseIf vt = "CM" Then
   
   rss.Open "select INVOICENO as VoucherID,'Counter Sale' as VoucherType,INVOICEDATE as VoucherDate,0 as VoucherNumber,netamount as Amount,SUBLEDGER as Particular from  casha where invoiceno = " + vno, con

ElseIf vt = "D" Then
   
   rss.Open "select DNN as VoucherID,'Debit Note' as VoucherType,DND as VoucherDate,0 as VoucherNumber,na as Amount,psld as Particular from  DNFA where dnn = " + vno, con

ElseIf vt = "C" Then
   
   rss.Open "select CNN as VoucherID,'Credit Note' as VoucherType,CND as VoucherDate,0 as VoucherNumber,na as Amount,psld as Particular from  CNF1A where Cnn = " + vno, con

End If

If (rss.EOF = False) Then


If vt = "V" Then

    con.Execute "insert into AuditTrail_Log(VoucherID,VoucherType,ActionType,VoucherDate,VoucherNumber,[Description],Amount,ReasionForEdit,UserName) " & _
    " values ('" & rss!VoucherID & "','" & rss!VoucherType & "','" & ActionType & "','" & Format(rss!voucherDATE, "MM/dd/yyyy") & "','" & rss!VOUCHERNUMBER & "','" & rss!Particular & "','" & rss!amount & "','" & ReasionForEdit & "','" & UserName & "')"

Else

    con.Execute "insert into AuditTrail_Log(VoucherID,VoucherType,ActionType,VoucherDate,VoucherNumber,[Description],Amount,ReasionForEdit,UserName) " & _
    " values ('" & rss!VoucherID & "','" & rss!VoucherType & "','" & ActionType & "','" & Format(rss!voucherDATE, "MM/dd/yyyy") & "','" & rss!VOUCHERNUMBER & "','" & rss!Particular & "','" & rss!amount & "','" & ReasionForEdit & "','" & UserName & "')"


End If


End If



End Function
Public Function ButtonPermissionNew(cmdSave As CommandButton, cmdDelete As CommandButton, cmdedit As CommandButton, frmName As String) As Boolean
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdedit.Enabled = False
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select [Save],[Delete],[Edit] from UsrePermission where userName='" & main.UserName & "' and TaskType='" & frmName & "'", coninfo, adOpenKeyset, adLockReadOnly
    
    If rs1.EOF = False Then
       
       If rs1(0) = "y" Then
          cmdSave.Enabled = True
       Else
          cmdSave.Enabled = False
       End If
       
       If rs1(1) = "y" Then
          cmdDelete.Enabled = True
       Else
          cmdDelete.Enabled = False
       End If
       
       If rs1(2) = "y" Then
          cmdedit.Enabled = True
       Else
          cmdedit.Enabled = False
       End If
       
    
    End If
End Function


Public Function fillDocument(code_ As String) As String

Dim DocDB As String
Dim lastYrs As String
lastYrs = ""

If (DateDiff("d", Now, SessionLastDate) <= 0) Then
   lastYrs = "current"
Else
   lastYrs = "last"
End If
 
FileDocument_Con



If (lastYrs = "last") Then

    dk = Mid(session_last1, 6)
    If Val(dk) >= 24 Then
     caf_ = "MOU-" & session_last1
    Else
     caf_ = "MOU"
    End If

Else

    dk = Mid(session, 6)
    If Val(dk) >= 24 Then
     caf_ = "MOU-" & session
    Else
     caf_ = "MOU"
    End If

End If

''dk = Mid(session, 6)
''If Val(dk) >= 24 Then
'' caf_ = "MOU-" & session
''Else
'' caf_ = "MOU"
''End If



Set rs1 = New ADODB.Recordset

rs1.Open "SELECT fname FROM PartyDocument where code='" & code_ & "' and LinkName ='" & caf_ & "'", con_doc1
If rs1.EOF = True Then
   fillDocument = "MOU Not Uploaded.."
Else

   If IsNull(rs1!fname) Then
      fillDocument = "MOU Not Uploaded.."
   End If
   
End If



End Function


Public Function checkPermission(st_ As String) As Boolean

Dim rs_ As New ADODB.Recordset

If st_ = "donnation" Then

   rs_.Open "select Permission from usrepermission where TaskType='mnuDonnation' and UserName='" & UserName & "'", coninfo
   If rs_.EOF = True Then
       checkPermission = False
   Else
       checkPermission = True
   End If
   
ElseIf st_ = "adj" Then

   rs_.Open "select Permission from usrepermission where TaskType='mnuAdjOp' and UserName='" & UserName & "'", coninfo
   If rs_.EOF = True Then
       checkPermission = False
   Else
       checkPermission = True
   End If
   
End If
      
       
End Function
Public Function RemoveEnterChar(strName As String) As String

Dim st_ As String
Dim a_strResult
st_ = ""
a_strResult = Split(strName, vbCr)



J = 1
k2 = 0
For k1 = 1 To 50
   s11 = Mid(strName, J + k1 - 1, 1)
   If (s11 <> vbCr And s11 <> vbLf) Then
      st_ = Mid(strName, J + k2, k1)
   Else
      k2 = k1 + 1
      k1 = 1
   End If
Next


End Function
Public Function MaxOrderNo(frmName As String) As String

Dim rs_ord As ADODB.Recordset
Set rs_ord = New ADODB.Recordset

rs_ord.Open "SELECT max(convert(int,ord_no)) FROM maxOrderNoQry where FirmName='" & frmName & "'", con
If Not IsNull(rs_ord(0)) Then
   MaxOrderNo = rs_ord(0) + 1
Else
   MaxOrderNo = 1
End If


End Function

Public Function GetCom() As String


Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetCom = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
    com_name = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
    
    
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String

     ' Get the user name minus any trailing spaces found in the name.
     ret = GetUserName(lpBuff, 25)
     com_user = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
     
     Sys_user_ = com_user

End Function
Sub FileDocument_Con()

Set con_doc1 = New ADODB.Connection
DocDB = "Database=chitraData_2223"

 

If LCase(server_) = "server" Then
   con_doc1.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & DocDB & "; UID=" & sql_user & "; PWD=" & sql_pass
Else
   con_doc1.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & DocDB & "; UID=""; PWD="""

End If

DoEvents
DoEvents


con_doc1.CursorLocation = adUseClient
If con_doc1.State = 1 Then con_doc1.close
con_doc1.Open

DoEvents
DoEvents

End Sub

Sub DSNNew()


Dim FSO As filesystemobject
Dim f As File
Dim txt As TextStream
Dim matter As String
Dim Total As String
Dim s(1, 2) As String
Set FSO = New filesystemobject
Dim ss, database_ As String
Dim rs10 As ADODB.Recordset

matter = ""

Dim op_system, Dusername, dstrpath As String
Dim systemName As String
systemName = ""

GetCom


Set rs10 = New ADODB.Recordset
rs10.Open "select * from UserDSN where UserName='" & com_name & "' and usrename_='" & com_user & "'", CCON

If rs10.EOF = False Then
   systemName = rs10!UserName
  dstrpath = rs10!Path '& "\chitradsn.dsn"
  Set txt = FSO.CreateTextFile("" & dstrpath, True)
End If



Set rs10 = New ADODB.Recordset
rs10.Open "select UID,WSID,[SERVER] from data", CCON
If rs10.EOF = False Then
   ss = "Server=" & rs10!server
   serverName_ = Mid(ss, 8)
   sql_pass = rs10!WSID & ""
   sql_user = rs10!UId & ""
End If




If session = "2016-17" Then
   database_ = "Database=chitraData_1617"
ElseIf session = "2015-16" Then
   database_ = "Database = chitraData"
ElseIf session = "2017-18" Then
   database_ = "Database=chitraData_1718"
ElseIf session = "2018-19" Then
   database_ = "Database=chitraData_1819"
ElseIf session = "2019-20" Then
   database_ = "Database=chitraData_1920"
ElseIf session = "2020-21" Then
   database_ = "Database=chitraData_2021"
ElseIf session = "2021-22" Then
   database_ = "Database=chitraData_2122"
ElseIf session = "2022-23" Then
   database_ = "Database=chitraData_2223"
ElseIf session = "2023-24" Then
   database_ = "Database=chitraData_2324"
ElseIf session = "2024-25" Then
   database_ = "Database=chitraData_2425"
End If


If server_ = "client" Then

    matter = matter & "[ODBC]" & vbNewLine
    matter = matter & "DRIVER=SQL Server" & vbNewLine
    matter = matter & "UId=" & sql_user & "" & vbNewLine
    matter = matter & "Trusted_Connection=Yes" & vbNewLine
    matter = matter & " & systemName & " & vbNewLine
    matter = matter & database_ & "" & vbNewLine
    matter = matter & "WSID=" & systemName & vbNewLine
    matter = matter & "APP=Microsoft Data Access Components" & vbNewLine
    matter = matter & ss & "" & vbNewLine
    txt.Write matter
    txt.close
    
Else

 
    matter = matter & "[ODBC]" & vbNewLine
    matter = matter & "DRIVER=SQL Server" & vbNewLine
    matter = matter & "UId=" & sql_user & "" & vbNewLine
    matter = matter & database_ & "" & vbNewLine
    matter = matter & "WSID=" & systemName & vbNewLine
    matter = matter & "APP=Microsoft® Windows® Operating System" & vbNewLine
    matter = matter & ss & "" & vbNewLine
    txt.Write matter
    txt.close
    

End If


''Set con = New ADODB.Connection
''constr = "filedsn=chitradsn;uid= " & sql_user  & ";pwd=" & sql_pass
''con.ConnectionString = constr
''con.CursorLocation = adUseClient




Set CONDSN = New ADODB.Connection
constr = "filedsn=chitradsn;uid=" & sql_user & ";pwd=" & sql_pass
CONDSN.ConnectionString = constr
If CONDSN.State = 1 Then CONDSN.close
  CONDSN.Open



End Sub
Sub createLog(user_ As String, no_ As String, vtype_ As String, desc_ As String, dates_ As String)
On Error GoTo aa_
con.Execute "insert into logtbl(UserName,No,Vtype,desc_,dates) values('" & user_ & "','" & no_ & "','" & vtype_ & "','" & desc_ & "','" & Format(dates_, "MM/dd/yyyy") & "')"
Exit Sub
aa_:
End Sub
Function fatchPrinter(k1 As Integer) As String

Dim s1 As String

If k1 = 1 Then
   s1 = "Inn_Printer"
ElseIf k1 = 2 Then
   s1 = "text_Printer"
ElseIf k1 = 3 Then
   s1 = "Exam_Printer"
ElseIf k1 = 4 Then
   s1 = "Supp_Printer"
ElseIf k1 = 5 Then
   s1 = "Title_Printer"
ElseIf k1 = 6 Then
   s1 = "cboPrinter6"
ElseIf k1 = 7 Then
   s1 = "cboPrinter7"
ElseIf k1 = 8 Then
   s1 = "cboPrinter8"
ElseIf k1 = 9 Then
   s1 = "cboPrinter9"
ElseIf k1 = 10 Then
   s1 = "cboPrinter10"
ElseIf k1 = 11 Then
   s1 = "cboPrinter11"
ElseIf k1 = 12 Then
   s1 = "cboPrinter12"
End If


fatchPrinter = s1
 

End Function
Function encrypt(txt As String) As String

Dim Estr As String
Dstr = txt

For I = 1 To Len(Dstr)
 Estr = Estr & Chr(Asc(Mid(Dstr, I, 1)) + 30)
Next I

encrypt = Estr
End Function
Function decrypt(txt As String) As String

Dim Estr As String, Dstr As String
Dstr = txt

For I = 1 To Len(Dstr)
 Estr = Estr & Chr(Asc(Mid(Dstr, I, 1)) + 30)
Next I

decrypt = Estr

End Function
'Function decrypt(txt As String) As String
'
'Dim Estr As String, Dstr As String
'Dstr = txt
'
'For I = 1 To Len(Dstr)
' Estr = Estr & Chr(Asc(Mid(Dstr, I, 1)) + 30)
'Next I
'
'decrypt = Estr
'
'End Function

Function ReturnDiscount(category As String, categorycode_ As String, gpcode As String) As Double


Dim rs11 As New ADODB.Recordset

If category = "C1" Then
          Set rs11 = con.Execute("select discountrate from DISCCATS_Sub where categorycode ='" + categorycode_ + "' and groupcode='" + gpcode + "'")
ElseIf category = "C2" Then
         Set rs11 = con.Execute("select discountrate from DISCCATS_Sub where categorycode ='" + categorycode_ + "' and groupcode='" + gpcode + "'")
ElseIf category = "C3" Then
         Set rs11 = con.Execute("select discountrate from DISCCATS_Sub where categorycode ='" + categorycode_ + "' and groupcode='" + gpcode + "'")
End If


If Not IsNull(rs11(0)) Then
If rs11.EOF = False Then
   ReturnDiscount = rs11(0)
Else
   ReturnDiscount = 0
End If
End If

End Function
Function ReturnBalanceSpQty(bcode As String, rep As String, ordno As String, ordDate As Date) As Double

Dim rs11 As New ADODB.Recordset
Dim rs12 As New ADODB.Recordset
Dim rs13 As New ADODB.Recordset
Dim qtyOrder, QtySp, qty As Integer
Dim last_ As String
Dim fDate_, tdate
qtyOrder = 0
QtySp = 0
qty = 0
last_ = "n"


fDate_ = "01/09/20" & Mid(session, 3, 2)
tdate_ = "30/09/20" & Right(session, 2)

If (ordDate >= fDate_ And ordDate <= tdate_) Then
   last_ = "n"
Else
   last_ = "y"
End If


If rs12.State = 1 Then rs12.close

If last_ = "n" Then
rs12.Open "select FromDate,ToDate from SpAllotmentQty where len(RepName)>0", con
Else
rs12.Open "select FromDate,ToDate from SpAllotmentQty where len(RepName)>0", con_LAST
End If

If rs12.EOF = False Then


If rs13.State = 1 Then rs13.close
If last_ = "n" Then
   rs13.Open "select top 1 * from OrderBookList where invoiceno='" & ordno & "' and  RepName='" & rep & "' and groupcode='BP' and (convert(datetime,invoiceDate,103)>=convert(datetime,'" & rs12!fromdate & "',103) and convert(datetime,invoiceDate,103)<=convert(datetime,'" & rs12!todate & "',103))", con
Else
   rs13.Open "select top 1 * from OrderBookList where invoiceno='" & ordno & "' and  RepName='" & rep & "' and groupcode='BP' and (convert(datetime,invoiceDate,103)>=convert(datetime,'" & rs12!fromdate & "',103) and convert(datetime,invoiceDate,103)<=convert(datetime,'" & rs12!todate & "',103))", con_LAST
End If

If rs13.EOF = False Then
   bcode = 0
End If


If rs11.State = 1 Then rs11.close
'rs11.Open "select sum(SpQty) as Qty from OrderBookList where RepName='" & rep & "' and groupcode='BP' and (convert(datetime,invoiceDate,103)>=convert(datetime,'" & rs12!FromDate & "',103) and convert(datetime,invoiceDate,103)<=convert(datetime,'" & rs12!toDate & "',103))", con
If last_ = "n" Then
   'Set rs11 = con.Execute("exec totalSpQty_Issued '" & rs12!FromDate & "','" & rs12!toDate & "','" & rep & "'")
   rs11.Open "select sum(Qty) from totalSpQty_Issued where RepName='" & rep & "' and (convert(datetime,invoiceDate,103)>=convert(datetime,'" & rs12!fromdate & "',103) and convert(datetime,invoiceDate,103)<=convert(datetime,'" & rs12!todate & "',103))", con
   
Else
   'Set rs11 = con_LAST.Execute("exec totalSpQty_Issued '" & rs12!FromDate & "','" & rs12!toDate & "','" & rep & "'")
   rs11.Open "select sum(Qty) from totalSpQty_Issued where RepName='" & rep & "' and (convert(datetime,invoiceDate,103)>=convert(datetime,'" & rs12!fromdate & "',103) and convert(datetime,invoiceDate,103)<=convert(datetime,'" & rs12!todate & "',103))", con_LAST
End If

If Not IsNull(rs11(0)) Then
   qtyOrder = rs11(0)
End If

If bcode <> "" Then
   qtyOrder = qtyOrder + Val(bcode)
End If



If rs11.State = 1 Then rs11.close
If last_ = "n" Then
   rs11.Open "select sum(Qty) from SpAllotmentQty where  repname='" & rep & "'", con
Else
   rs11.Open "select sum(Qty) from SpAllotmentQty where  repname='" & rep & "'", con_LAST
End If
If Not IsNull(rs11(0)) Then
   QtySp = rs11(0)
End If


If rs11.State = 1 Then rs11.close
If last_ = "n" Then
'Set rs11 = con.Execute("exec SpecimenAlotment '" & rs12!FromDate & "','" & rs12!toDate & "','" & rep & "'")
rs11.Open "select sum(Qty) from TotalSpReturnRepWise where  agentname='" & rep & "'", con
Else
'Set rs11 = con_LAST.Execute("exec SpecimenAlotment '" & rs12!FromDate & "','" & rs12!toDate & "','" & rep & "'")
rs11.Open "select sum(Qty) from TotalSpReturnRepWise where  agentname='" & rep & "'", con_LAST

End If
If Not IsNull(rs11(0)) Then
   qtyOrder = qtyOrder - rs11(0)
End If
        

qty = QtySp - qtyOrder



End If


ReturnBalanceSpQty = qty


End Function

Function ReturnDiscountNew(bcode As String, pname As String, scid As String) As Double


Dim rs11 As New ADODB.Recordset
Dim rs12 As New ADODB.Recordset
Dim category As String
Dim subgp, gpcode As String

Dim dis As Double
dis = 0
subgp = "n"

ReturnDiscountNew = 0

str_ = "SELECT BOOKCODE,BOOKNAME,GROUPCODE,DISCOUNT,Category,GROUPCODE_sub FROM BookQry_ where bookcode='" & bcode & "'"
str10 = "select DISCATEGORY,Category2,Category3 from SLEDGER  where SUBLEDGER='" & pname & "'"


If rs12.State = 1 Then rs12.close
rs12.Open str_, con
If rs12.EOF = False Then
   ReturnDiscountNew = rs12!discount
   category = rs12!category
   
   
   If (Not IsNull(rs12!GROUPCODE_sub) And rs12!GROUPCODE_sub <> "") Then
           subgp = "y"
           gpcode = rs12!GROUPCODE_sub
   Else
      subgp = "n"
      gpcode = rs12!groupcode
   End If
   
End If

If rs12.State = 1 Then rs12.close
rs12.Open str10, con
If rs12.EOF = False Then
   If category = "C1" Then
      pcategory = rs12!DISCATEGORY
   ElseIf category = "C2" Then
      pcategory = rs12!category2
   ElseIf category = "C3" Then
      pcategory = rs12!Category3
   End If
End If


If category = "C1" Then
   If subgp = "n" Then
      Set rs11 = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   Else
      Set rs11 = con.Execute("select discountrate from DISCCATS_Sub where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   End If
   
ElseIf category = "C2" Then
   If subgp = "n" Then
      Set rs11 = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   Else
      Set rs11 = con.Execute("select discountrate from DISCCATS_Sub where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   End If
ElseIf category = "C3" Then
   If subgp = "n" Then
     Set rs11 = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   Else
     Set rs11 = con.Execute("select discountrate from DISCCATS_Sub where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   End If
   
End If


If Not IsNull(rs11(0)) Then
If rs11.EOF = False Then
   ReturnDiscountNew = rs11(0)
End If
End If





''old code
''   If scid = "" Then
''
''        If rs11.State = 1 Then rs11.close
''        Set rs11 = con.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Mid(pname, 1, 5) & "', '" & bcode & "',''")
''        If rs11.EOF = False Then
''           ReturnDiscountNew = rs11(0)
''        End If
''
''     Else
''
''        If rs11.State = 1 Then rs11.close
''        Set rs11 = con.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Mid(pname, 1, 5) & "', '" & bcode & "', '" & scid & "' ")
''        If rs11.EOF = False Then
''           ReturnDiscountNew = rs11(0)
''        End If
''
''     End If

    
    

If (DateDiff("d", Now, SessionLastDate) <= 0) Then
    
    
    If scid = "" Then
    
        If rs11.State = 1 Then rs11.close
        Set rs11 = con.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Mid(pname, 1, 5) & "', '" & bcode & "',''")
        If rs11.EOF = False Then
           ReturnDiscountNew = rs11(0)
        End If
        
     Else
     
        If rs11.State = 1 Then rs11.close
        Set rs11 = con.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Mid(pname, 1, 5) & "', '" & bcode & "', '" & scid & "' ")
        If rs11.EOF = False Then
           ReturnDiscountNew = rs11(0)
        End If
        
     End If
 
 
 Else

     If scid = "" Then

        If rs11.State = 1 Then rs11.close
        Set rs11 = con_LAST.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Mid(pname, 1, 5) & "', '" & bcode & "',''")
        If rs11.EOF = False Then
           ReturnDiscountNew = rs11(0)
        End If

     Else

        If rs11.State = 1 Then rs11.close
        Set rs11 = con_LAST.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Mid(pname, 1, 5) & "', '" & bcode & "', '" & scid & "' ")
        If rs11.EOF = False Then
           ReturnDiscountNew = rs11(0)
        End If

     End If


 End If
 
 
 
 
 
 
 


End Function
Function ReturnDiscountNew_Return(bcode As String, pname As String, scid As String) As Double


Dim rs11 As New ADODB.Recordset
Dim rs12 As New ADODB.Recordset
Dim category As String
Dim subgp, gpcode As String

Dim dis As Double
dis = 0
subgp = "n"



ReturnDiscountNew_Return = 0

str_ = "SELECT BOOKCODE,BOOKNAME,GROUPCODE,DISCOUNT,Category,GROUPCODE_sub FROM BookQry_ where bookcode='" & bcode & "'"
str10 = "select DISCATEGORY,Category2,Category3 from SLEDGER  where SUBLEDGER='" & pname & "'"


If rs12.State = 1 Then rs12.close
rs12.Open str_, con
If rs12.EOF = False Then
   ReturnDiscountNew_Return = rs12!discount
   category = rs12!category
   
   
   If (Not IsNull(rs12!GROUPCODE_sub) And rs12!GROUPCODE_sub <> "") Then
           subgp = "y"
           gpcode = rs12!GROUPCODE_sub
   Else
      subgp = "n"
      gpcode = rs12!groupcode
   End If
   
End If

If rs12.State = 1 Then rs12.close
rs12.Open str10, con
If rs12.EOF = False Then
   If category = "C1" Then
      pcategory = rs12!DISCATEGORY
   ElseIf category = "C2" Then
      pcategory = rs12!category2
   ElseIf category = "C3" Then
      pcategory = rs12!Category3
   End If
End If


If category = "C1" Then
   If subgp = "n" Then
      Set rs11 = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   Else
      Set rs11 = con.Execute("select discountrate from DISCCATS_Sub where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   End If
   
ElseIf category = "C2" Then
   If subgp = "n" Then
      Set rs11 = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   Else
      Set rs11 = con.Execute("select discountrate from DISCCATS_Sub where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   End If
ElseIf category = "C3" Then
   If subgp = "n" Then
     Set rs11 = con.Execute("select discountrate from DISCCATS where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   Else
     Set rs11 = con.Execute("select discountrate from DISCCATS_Sub where " & stringyear & " and categorycode ='" + pcategory + "' and groupcode='" + gpcode + "'")
   End If
   
End If


If Not IsNull(rs11(0)) Then
If rs11.EOF = False Then
   ReturnDiscountNew_Return = rs11(0)
End If
End If



''old code
''   If scid = "" Then
''
''        If rs11.State = 1 Then rs11.close
''        Set rs11 = con.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Mid(pname, 1, 5) & "', '" & bcode & "',''")
''        If rs11.EOF = False Then
''           ReturnDiscountNew = rs11(0)
''        End If
''
''     Else
''
''        If rs11.State = 1 Then rs11.close
''        Set rs11 = con.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Mid(pname, 1, 5) & "', '" & bcode & "', '" & scid & "' ")
''        If rs11.EOF = False Then
''           ReturnDiscountNew = rs11(0)
''        End If
''
''     End If

    
    
If (DateDiff("d", Now, SessionLastDate) <= 0) Then
        
        
    
    If rs11.State = 1 Then rs11.close
    rs11.Open "select sername from books where bookcode='" & bcode & "'", con
    If rs11.EOF = False Then
      sername = rs11(0)
    End If
   
   
    If scid = "" Then
     
        If rs11.State = 1 Then rs11.close
        Set rs11 = con.Execute("select LastYrs_Discount from SeriesWiseDiscountQry_New where Party= '" & pname & "' and SeriesName = '" & sername & "'")
        If rs11.EOF = False Then
            ReturnDiscountNew_Return = rs11(0)
        End If
        
     Else
     
        If rs11.State = 1 Then rs11.close
        Set rs11 = con.Execute("select LastYrs_Discount from SeriesWiseDiscountQry_New where Party= '" & pname & "' and SeriesName = '" & sername & "' and ScId='" & scid & "'")
        If rs11.EOF = False Then
           ReturnDiscountNew_Return = rs11(0)
        End If
        
     End If
 
 

 End If
 
 
 

End Function
Function smsSend(invoiceNo As String, mob As String, inv As String)
     
On Error Resume Next
     
Screen.MousePointer = vbHourglass

    Dim myURL As String
    Dim message As String
    Dim winHttpReq As Object

    message = "Dear Patron%nYour order has been dispatched. Click here for details%n http://blueprinteducation.co.in/inv.php?invId=" & inv & " %n %nThanks%nBlueprint Education"
    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    myURL = "https://api.textlocal.in/send/?"
    winHttpReq.Open "POST", myURL, False
    winHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    winHttpReq.Send ("username=nitin@blueprinteducation.co.in&hash=43cfafad3b06817ebf82e26befcab56da90fa3b6&sender=BPEDUC&numbers=" & mob & "&message=" & message)
    SendSMS = winHttpReq.responseText
    
    If InStr(SendSMS, "status") > 0 Then
       If Mid(inv, 1, 2) = "IN" Then
          con.Execute "update invoicea set SMSSend='y' where invoiceno=" & invoiceNo & ""
       Else
          con.Execute "update INVOICEA_sp set SMSSend='y' where invoiceno=" & invoiceNo & ""
       End If
   End If


Screen.MousePointer = vbDefault

End Function
Function value_() As String

''''======================

clave = ""
all_chars = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "S", "T", "U", "V", "W", "X", "Y", "Z")
Randomize

For I = 1 To 3
   random_index = Int(Rnd() * 25)
   clave = clave & all_chars(random_index)
Next

 value_ = clave
    
    
End Function
Public Sub createExcel(ByRef sqlQry As String, ByRef rtype_ As String)

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String
Dim soldTillDate As Long
Dim from_date, last_date As Date
Dim db As String



If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add

intSheets = xlBook.Worksheets.Count

Set xlSheetLast = xlBook.Worksheets(intSheets)
Set xlSheet = xlBook.Worksheets.Add



Dim c, r, r1 As Long
Dim Q1, q2, J As Integer

c = 1
r = 1

xl.Columns("A:H").ColumnWidth = 15
J = 2


If rs_1.State = 1 Then rs_1.close
rs_1.Open sqlQry, con, adOpenDynamic, adLockOptimistic
For r = 0 To rs_1.Fields.Count - 1
    xlSheet.Cells(1, c).value = rs_1.Fields(r).Name
    c = c + 1
Next

Dim b1 As Boolean
b1 = True

If rtype_ = "1" Then

For r = 1 To rs_1.RecordCount

    If rs_1.EOF = False Then
        xlSheet.Cells(r + 1, 1).value = rs_1.Fields(0).value
        xlSheet.Cells(r + 1, 2).value = rs_1.Fields(1).value
        xlSheet.Cells(r + 1, 3).value = rs_1.Fields(2).value
        xlSheet.Cells(r + 1, 4).value = rs_1.Fields(3).value
        xlSheet.Cells(r + 1, 5).value = rs_1.Fields(4).value
        xlSheet.Cells(r + 1, 6).value = rs_1.Fields(5).value
        xlSheet.Cells(r + 1, 7).value = rs_1.Fields(6).value
        xlSheet.Cells(r + 1, 8).value = rs_1.Fields(7).value
        
        If r = 1 Then
           xlSheet.Cells(r + 1, 9).value = rs_1.Fields(8).value
           xlSheet.Cells(r + 1, 10).value = rs_1.Fields(9).value
        ElseIf ss10 <> rs_1.Fields(5).value Then
           xlSheet.Cells(r + 1, 9).value = rs_1.Fields(8).value
           xlSheet.Cells(r + 1, 10).value = rs_1.Fields(9).value
        Else
           xlSheet.Cells(r + 1, 9).value = 0
           xlSheet.Cells(r + 1, 10).value = 0
        End If
            
        ss10 = rs_1.Fields(5).value
        
        
        r1 = r1 + 1
    End If
    rs_1.MoveNext
    
    
Next


ElseIf rtype_ = 2 Then


For r = 1 To rs_1.RecordCount
    If rs_1.EOF = False Then
        xlSheet.Cells(r + 1, 1).value = rs_1.Fields(0).value
        xlSheet.Cells(r + 1, 2).value = rs_1.Fields(1).value
        xlSheet.Cells(r + 1, 3).value = rs_1.Fields(2).value
        xlSheet.Cells(r + 1, 4).value = rs_1.Fields(3).value
        xlSheet.Cells(r + 1, 5).value = rs_1.Fields(4).value
        xlSheet.Cells(r + 1, 6).value = rs_1.Fields(5).value
        xlSheet.Cells(r + 1, 7).value = rs_1.Fields(6).value
        
        

        
        
        r1 = r1 + 1
    End If
    rs_1.MoveNext
    
    
Next








Else

For r = 1 To rs_1.RecordCount - 1
    If rs_1.EOF = False Then
     For I = 1 To c - 1
        xlSheet.Cells(r + 1, I).value = rs_1.Fields(I - 1).value
        r1 = r1 + 1
     Next
    End If
    rs_1.MoveNext
Next


End If
        



Screen.MousePointer = vbDefault


Exit Sub

Screen.MousePointer = vbDefault

err:

MsgBox err.DESCRIPTION



End Sub
Function checkData_ForThisNumber(tblName As String, I_NO As Long, i_dt As Date) As Boolean

    
Dim tRS1 As New ADODB.Recordset
Dim trs2 As New ADODB.Recordset
If trs2.State = 1 Then trs2.close
trs2.Open "Select top 2 invoiceno as cn from " & tblName & "", con, adOpenDynamic, adLockOptimistic
If trs2.RecordCount <= 0 Then
   
   Exit Function
Else
    If tRS1.State = 1 Then tRS1.close
    tRS1.Open "Select top 10 min(invoiceno) as mid,invoicedate from  " & tblName & "  group by invoiceno,invoiceDate", con, adOpenDynamic, adLockOptimistic
    
    If tRS1.RecordCount > 0 Then
            If CDate(i_dt) <= tRS1!invoiceDate Then
                If CDate(i_dt) <> tRS1!invoiceDate Then
                     If Month(CDate(i_dt)) <> 4 And Day(CDate(i_dt)) <> 1 Then
                        MsgBox "Please select valid Invoice No. for this date.."
                        checkData_ForThisNumber = True
                        Exit Function
                     Else
                         If tRS1!Mid <> 1 Then
                            If Val(I_NO) >= tRS1!Mid Then
                              checkData_ForThisNumber = True
                            Exit Function
                        End If
                     End If
               End If
           End If
     End If
End If
End If
        
    If trs2.State = 1 Then trs2.close
    trs2.Open "Select max(invoiceno) as mid from " & tblName & " where  invoicedate <= convert(smalldatetime,'" & i_dt & "',103)-1", con, adOpenDynamic, adLockOptimistic
    If trs2.RecordCount > 0 Then
        If IsNull(trs2!Mid) <> True Then
            If Val(I_NO) >= trs2!Mid Then
               If tRS1.State = 1 Then tRS1.close
               tRS1.Open "Select  min(InvoiceNo)as m2 from " & tblName & " where invoicedate >= convert(smalldatetime,'" & i_dt & "',103)+1", con, adOpenDynamic, adLockOptimistic
               If tRS1.RecordCount > 0 Then
                  If IsNull(tRS1!m2) <> True Then
                     If Val(I_NO) <= tRS1!m2 Then

                     Else
                         checkData_ForThisNumber = True
                     End If
                  End If
               End If

            Else
                    checkData_ForThisNumber = True
            End If
     End If
    End If

End Function
Public Function DebitFromRepSaleNew()
   
   Dim rs10 As New ADODB.Recordset
   
   debitForAgnNew = ""
   
   Set rs10 = New ADODB.Recordset
   rs10.Open "select gledger,DebitFromRepSale from GLEDGER where DebitFromRepSale=1", con
   While rs10.EOF = False
     
     If debitForAgnNew = "" Then
        debitForAgnNew = "PGLD='" & rs10(0) & "'"
     Else
        debitForAgnNew = debitForAgnNew & " or PGLD='" & rs10(0) & "'"
     End If
     
     rs10.MoveNext
   Wend
   
   If debitForAgnNew <> "" Then
      DebitFromRepSaleNew = "(" & debitForAgnNew & ")"
   End If
   
End Function
Public Function DebitFromRepSale()
   
   Dim rs10 As New ADODB.Recordset
   
   debitForAgn = ""
   
   Set rs10 = New ADODB.Recordset
   rs10.Open "select gledger,DebitFromRepSale from GLEDGER where DebitFromRepSale=1", con
   While rs10.EOF = False
     
     If debitForAgn = "" Then
        debitForAgn = "GLD='" & rs10(0) & "'"
     Else
        debitForAgn = debitForAgn & " or GLD='" & rs10(0) & "'"
     End If
     
     rs10.MoveNext
   Wend
   
   If debitForAgn <> "" Then
      debitForAgn = "(" & debitForAgn & ")"
   End If
   
End Function
Public Sub financialyear()


Set RS = New ADODB.Recordset
RS.Open "select fromDate,toDate,NotCreated,fromDateSRet,toDateSRet from turnOverDis where fyear='" & session & "'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
     financialyear_Fdate = RS!fromdate
     financialyear_Tdate = RS!todate
     
     financialyear_Fdate_SaleRet = RS!fromDateSRet
     financialyear_Tdate_SaleRet = RS!toDateSRet
     
  End If
End If


Set RS = New ADODB.Recordset
RS.Open "select fromDate,toDate,NotCreated,DataBase,toDateSRet from turnOverDis where Current_Next='next'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
    financialyear_Tdate = RS!todate
    financialyear_Tdate_SaleRet = RS!toDateSRet
  End If
End If

End Sub
Public Sub addmaster_addSingleData(str_ As String, rd As String, Optional findwhere As String)

Dim rs_sql As New ADODB.Recordset
Dim rs_mdb As New ADODB.Recordset

Select Case str_

Case "Books"

    If rs_sql.State = 1 Then rs_sql.close
    rs_sql.Open "select * from books where bookcode='" & rd & "' AND " & stringyear, con, adOpenDynamic, adLockOptimistic
    If rs_mdb.State = 1 Then rs_mdb.close
    rs_mdb.Open "select * from books where bookcode='" & rd & "'", CCON, adOpenDynamic, adLockOptimistic
    If rs_mdb.EOF = True Then
        rs_mdb.AddNew
    End If
    rs_mdb!Bookcode = rs_sql!Bookcode
    rs_mdb!Bookname = rs_sql!Bookname
    rs_mdb!groupcode = rs_sql!groupcode
    rs_mdb!rate = rs_sql!rate
    rs_mdb!discount = rs_sql!discount
    rs_mdb!fyear = rs_sql!fyear
    rs_mdb!setupid = rs_sql!setupid
    rs_mdb.update
    
Case "sledger"

If rs_sql.State = 1 Then rs_sql.close
rs_sql.Open "select * from sledger where SUBLEDGER='" & rd & "' AND " & stringyear, con, adOpenDynamic, adLockOptimistic

If rs_mdb.State = 1 Then rs_mdb.close
rs_mdb.Open "select * from sledger where SUBLEDGER='" & findwhere & "'", CCON, adOpenDynamic, adLockOptimistic
If rs_mdb.EOF = True Then
    rs_mdb.AddNew
End If
    rs_mdb!gledger = rs_sql!gledger
    rs_mdb!subledger = rs_sql!subledger
    rs_mdb!party = rs_sql!party
    rs_mdb!Code = rs_sql!Code
    rs_mdb!YEAROPENING = rs_sql!YEAROPENING
    rs_mdb!DESCFORINVOICE = rs_sql!DESCFORINVOICE
    rs_mdb!fyear = rs_sql!fyear
    rs_mdb!setupid = rs_sql!setupid
    rs_mdb.update
    rs_sql.MoveNext

End Select


End Sub
Public Sub addmaster()


Screen.MousePointer = vbHourglass

Dim tbl_src, tbl_targ As Integer
Dim rs_sql As New ADODB.Recordset
Dim rs_mdb As New ADODB.Recordset

'transportmaster
'==============================================================
'''''''''''''''''''''''''''''''''''''''''''''  transportmaster -------------------------------------------------------------
'==============================================================
'==============================================================

CCON.Execute "delete from transportmaster where " & stringyear
If rs_sql.State = 1 Then rs_sql.close
rs_sql.Open "select * from transportmaster where " & stringyear, con, adOpenDynamic, adLockOptimistic

If rs_mdb.State = 1 Then rs_mdb.close
rs_mdb.Open "select * from transportmaster", CCON, adOpenDynamic, adLockOptimistic

While rs_sql.EOF = False
    rs_mdb.AddNew
    rs_mdb!transportname = rs_sql!transportname
    rs_mdb!fyear = rs_sql!fyear
    rs_mdb!setupid = rs_sql!setupid
    rs_mdb.update
    rs_sql.MoveNext
Wend

'==============================================================
'==============================================================
'''''''''''''''''''''''''''''''''''''''''''''  book -------------------------------------------------------------
'==============================================================
'==============================================================

CCON.Execute "delete from books where " & stringyear

If rs_sql.State = 1 Then rs_sql.close
rs_sql.Open "select * from books where " & stringyear, con, adOpenDynamic, adLockOptimistic

If rs_mdb.State = 1 Then rs_mdb.close
rs_mdb.Open "select * from books", CCON, adOpenDynamic, adLockOptimistic

While rs_sql.EOF = False
    rs_mdb.AddNew
    rs_mdb!Bookcode = rs_sql!Bookcode
    rs_mdb!Bookname = rs_sql!Bookname
    rs_mdb!groupcode = rs_sql!groupcode
    rs_mdb!rate = rs_sql!rate
    rs_mdb!discount = rs_sql!discount
    rs_mdb!fyear = rs_sql!fyear
    rs_mdb!setupid = rs_sql!setupid
    rs_mdb.update
    rs_sql.MoveNext
Wend
'===================================================================================================================
'''''''''''''''''''''''''''''''''''''''''''''  sledger -------------------------------------------------------------
'===================================================================================================================
'===================================================================================================================

CCON.Execute "delete from sledger where " & stringyear
If rs_sql.State = 1 Then rs_sql.close
rs_sql.Open "select * from sledger where " & stringyear, con, adOpenDynamic, adLockOptimistic

If rs_mdb.State = 1 Then rs_mdb.close
rs_mdb.Open "select * from sledger", CCON, adOpenDynamic, adLockOptimistic

While rs_sql.EOF = False
    rs_mdb.AddNew
    rs_mdb!gledger = rs_sql!gledger
    rs_mdb!subledger = rs_sql!subledger
    rs_mdb!party = rs_sql!party
    rs_mdb!Code = rs_sql!Code
    rs_mdb!YEAROPENING = rs_sql!YEAROPENING
    rs_mdb!DESCFORINVOICE = rs_sql!DESCFORINVOICE
    rs_mdb!fyear = rs_sql!fyear
    rs_mdb!setupid = rs_sql!setupid
    rs_mdb.update
    rs_sql.MoveNext
Wend
'-------------------------------------------------------------------------------
'===============================================================================
'''''''''''''''''''''''''''''''''''''''''''''  school -------------------------------------------------------------
'===============================================================================
'===============================================================================

CCON.Execute "delete from school where " & stringyear
If rs_sql.State = 1 Then rs_sql.close
rs_sql.Open "select distinct aname,city,scname,fyear,setupid  from info where " & stringyear, con, adOpenDynamic, adLockOptimistic

If rs_mdb.State = 1 Then rs_mdb.close
rs_mdb.Open "select * from school", CCON, adOpenDynamic, adLockOptimistic

While rs_sql.EOF = False
    rs_mdb.AddNew
    rs_mdb!scname = rs_sql!scname
    rs_mdb!agentname = rs_sql!aname
    rs_mdb!city = rs_sql!city
    rs_mdb!fyear = rs_sql!fyear
    rs_mdb!setupid = rs_sql!setupid
    rs_mdb.update
    rs_sql.MoveNext
Wend

'-------------------------------------------------------------------------------
'===============================================================================
'''''''''''''''''''''''''''''''''''''''''''''  school -------------------------------------------------------------
'===============================================================================
'===============================================================================

CCON.Execute "delete from BookMaster where " & stringyear
If rs_sql.State = 1 Then rs_sql.close
rs_sql.Open "select distinct BookNo,Book,class,bookfont,fyear,setupid  from BookMaster where " & stringyear, con, adOpenDynamic, adLockOptimistic

If rs_mdb.State = 1 Then rs_mdb.close
rs_mdb.Open "select * from BookMaster", CCON, adOpenDynamic, adLockOptimistic
While rs_sql.EOF = False
    
    rs_mdb.AddNew
    rs_mdb!bookNo = rs_sql!bookNo
    rs_mdb!Book = rs_sql!Book
    rs_mdb!Class = rs_sql!Class
    rs_mdb!bookfont = rs_sql!bookfont
    rs_mdb!fyear = rs_sql!fyear
    rs_mdb!setupid = rs_sql!setupid
    rs_mdb.update
    rs_sql.MoveNext

Wend

'-------------------------------------------------------------------------------
'===============================================================================
'''''''''''''''''''''''''''''''''''''''''''''  college -------------------------------------------------------------
'===============================================================================
'===============================================================================
CCON.Execute "delete from College where " & stringyear
If rs_sql.State = 1 Then rs_sql.close
rs_sql.Open "select distinct CollegeID,College,city,district,states,fyear,setupid  from College where " & stringyear, con, adOpenDynamic, adLockOptimistic
If rs_mdb.State = 1 Then rs_mdb.close
rs_mdb.Open "select * from College", CCON, adOpenDynamic, adLockOptimistic
While rs_sql.EOF = False
    rs_mdb.AddNew
        rs_mdb!collegeid = rs_sql!collegeid
            rs_mdb!college = rs_sql!college
                rs_mdb!city = rs_sql!city
                    rs_mdb!District = rs_sql!District
                rs_mdb!states = rs_sql!states
            rs_mdb!fyear = rs_sql!fyear
        rs_mdb!setupid = rs_sql!setupid
    rs_mdb.update
rs_sql.MoveNext
Wend


'Update Free-------------------------------------------------------------------------------
''Dim kk As New ADODB.Recordset
''
''CON.Execute "delete from INVOICEBSP_Free"
''CON.Execute "delete from INVOICEB_Free"
''CON.Execute "delete from BookStock_free"
''
''
''If RS.State = 1 Then RS.close
''RS.Open "select INVOICENO,INVOICEDATE,Genledger,Bookcode,SUBLEDGER,agentname,Godown,quantity from invoiceBQry order by INVOICENO", CON
''While RS.EOF = False
''
''If kk.State = 1 Then kk.close
''kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,apply from KitQry where kitcode='" & RS!Bookcode & "'", CON
''While kk.EOF = False
''
'' If kk!Apply = "y" Then
''    CON.Execute "insert into INVOICEB_Free(INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
''    " values('" & RS!INVOICENO & "','" & Format(RS.Fields("INVOICEDATE").value, "MM/dd/yyyy") & "','" & RS!Genledger & "','" & RS!SUBLEDGER & "','" & kk!Bookcode & "','" & (kk!qty * RS!quantity) & "','" & kk!rate & "','" & RS!agentname & "','" & setupid & "','" & session & "','" & RS!Godown & "')"
'' End If
''
''    kk.MoveNext
''Wend
''
''RS.MoveNext
''Wend
''
''''INVOICEBSP_Free
''
''If RS.State = 1 Then RS.close
''RS.Open "select INVOICENO,INVOICEDATE,Genledger,Bookcode,SUBLEDGER,agentname,Godown,quantity from invoiceSPBQry order by INVOICENO", CON
''While RS.EOF = False
''
''If kk.State = 1 Then kk.close
''kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,apply from KitQry where kitcode='" & RS!Bookcode & "'", CON
''While kk.EOF = False
''
''If kk!Apply = "y" Then
''    CON.Execute "insert into INVOICEBSP_Free(INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,BOOKCODE,QUANTITY,RATE,agentname,setupid,Fyear,Godown) " & _
''    " values('" & RS!INVOICENO & "','" & Format(RS.Fields("INVOICEDATE").value, "MM/dd/yyyy") & "','" & RS!Genledger & "','" & RS!SUBLEDGER & "','" & kk!Bookcode & "','" & (kk!qty * RS!quantity) & "','" & kk!rate & "','" & RS!agentname & "','" & setupid & "','" & session & "','" & RS!Godown & "')"
''End If
''
''    kk.MoveNext
''Wend
''
''RS.MoveNext
''Wend
''
''
''
''
''
''
''''
''''StockIssue_Free
''
''If RS.State = 1 Then RS.close
''RS.Open "SELECT EntryNo,Dates,BOOKCODE,Qty,Issue_Receive,Godown_Out FROM BookStock where Issue_Receive='Issue' order by EntryNo", CON
''While RS.EOF = False
''If kk.State = 1 Then kk.close
''kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,apply from KitQry where kitcode='" & RS!Bookcode & "'", CON
''While kk.EOF = False
''
'' If kk!Apply = "y" Then
''    CON.Execute "insert into BookStock_free(EntryNo,Dates,BOOKCODE,Qty,setupid,Fyear,Godown,Issue_Receive) " & _
''    " values('" & RS!EntryNo & "','" & Format(RS.Fields("Dates").value, "MM/dd/yyyy") & "','" & kk!Bookcode & "','" & (kk!qty * RS!qty) & "','" & setupid & "','" & session & "','" & RS!Godown_Out & "','" & RS!Issue_Receive & "')"
'' End If
''
''    kk.MoveNext
''Wend
''
''RS.MoveNext
''Wend


''''''StockReceive_Free
''''
''''If RS.State = 1 Then RS.close
''''RS.Open "SELECT EntryNo,Dates,BOOKCODE,Qty,Issue_Receive,Godown_In FROM BookStock where Issue_Receive='Receive' order by EntryNo", CON
''''While RS.EOF = False
''''If kk.State = 1 Then kk.close
''''kk.Open "select BOOKNAME,NoPrintDesc,Qty,bookcode,rate,apply from KitQry where kitcode='" & RS!Bookcode & "'", CON
''''While kk.EOF = False
''''
'''' If kk!Apply = "y" Then
''''    CON.Execute "insert into BookStock_free(EntryNo,Dates,BOOKCODE,Qty,setupid,Fyear,Godown,Issue_Receive) " & _
''''    " values('" & RS!EntryNo & "','" & Format(RS.Fields("Dates").value, "MM/dd/yyyy") & "','" & kk!Bookcode & "','" & (kk!qty * RS!qty) & "','" & setupid & "','" & session & "','" & RS!Godown_In & "','" & RS!Issue_Receive & "')"
'''' End If
''''
''''    kk.MoveNext
''''Wend
''''
''''RS.MoveNext
''''Wend
''''


''
''
''


Screen.MousePointer = vbDefault

End Sub
Public Sub clear_grid(g As MSFlexGrid)

For d1 = 1 To g.rows - 1
 g.TextMatrix(d1, 0) = ""
 g.TextMatrix(d1, 1) = ""
 g.TextMatrix(d1, 2) = ""
 g.TextMatrix(d1, 3) = ""
 g.TextMatrix(d1, 4) = ""
 g.TextMatrix(d1, 5) = ""
 g.TextMatrix(d1, 6) = ""
 g.TextMatrix(d1, 7) = ""
 g.TextMatrix(d1, 8) = ""
Next



End Sub
Public Function checkPacking(billno As Long, category As String) As Long
   
If rs1.State = 1 Then rs1.close
rs1.Open "select sum(Qty) from PackingQry where Category='" & category & "' and billno=" & billno & "", con
If Not IsNull(rs1(0)) Then
   checkPacking = rs1(0)
Else
   checkPacking = 0
End If
     
End Function
Public Function ReturnBookDesc(bcode As String) As String
Dim bkdesc As String

If rs1.State = 1 Then rs1.close
rs1.Open "select BOOKNAME,NoPrintDesc from KitQry where kitcode='" & bcode & "'", con
bkdesc = ""
While rs1.EOF = False
 If bkdesc = "" Then
   If rs1!NoPrintDesc = False Then
    bkdesc = rs1!Bookname
   End If
 Else
   If rs1!NoPrintDesc = False Then
    bkdesc = bkdesc & "," & rs1!Bookname
   End If
 End If
 rs1.MoveNext
Wend

If bkdesc <> "" Then
   ReturnBookDesc = "(" & bkdesc & ")"
End If

End Function
Public Function Grid_Validation(J As Integer) As Boolean
 
 Dim a As Boolean
 If J >= 48 And J <= 57 Or J = 8 Or J = 13 Then
   Grid_Validation = True
 Else
   Grid_Validation = False
 End If

End Function
Sub RefData(f As Form)

On Error Resume Next
Dim o As Object
For Each o In f
If TypeOf o Is textbox Or TypeOf o Is ComboBox Then
   o.text = ""
End If
Next

End Sub
Public Function Cal_ReamAndSheet(ByVal rm_ As Integer, ByVal st_ As Integer)
    
Dim cal_ream
If st_ > 499 Then
   cal_ream = Int(st_ / 500)
   sheet_tot = st_ - cal_ream * 500
   ream_tot = rm_ + cal_ream
Else
   ream_tot = rm_
   sheet_tot = st_
End If

End Function
Sub UpdateDisPatchReg1(inv As Integer, invDt As String, pname As String, station As String, bundle As String, transport As String, marka As String, bilty As String, grData As String, frieght As String, tlb As String)
    
    Dim rs_dis As New ADODB.Recordset
    Dim RS As New Recordset
    If RS.State = 1 Then RS.close
    RS.Open "Select max(" & "SNO" & ") from " & tlb, con, adOpenKeyset, adLockReadOnly
    If IsNull(RS(0)) Then
        kk = 1
    Else
        kk = Val(RS(0)) + 1
    End If
    RS.close
     
    If rs_dis.State = 1 Then rs_dis.close
    rs_dis.Open "select * from " & tlb & " where cno=" & inv & "", con, adOpenDynamic, adLockOptimistic
    If rs_dis.EOF = True Then
       rs_dis.AddNew
       rs_dis!sno = kk
       rs_dis!Date = invDt
       rs_dis!Particulars = pname
       rs_dis!BDL = station
       rs_dis!wt = bundle
       rs_dis!freight = transport
       rs_dis!rr = marka
       rs_dis!gr = Val(bilty)
       If IsDate(grData) Then
       rs_dis!GR_DT = grData
       End If
       rs_dis!Freight_Paid = frieght
       rs_dis!CMNo = inv
       rs_dis.update
    Else
       rs_dis!Date = invDt
       rs_dis!Particulars = pname
       rs_dis!BDL = station
       rs_dis!wt = bundle
       rs_dis!freight = transport
       rs_dis!rr = marka
       rs_dis!gr = Val(bilty)
       If IsDate(grData) Then
       rs_dis!GR_DT = grData
       End If
       rs_dis!Freight_Paid = frieght
       rs_dis.update
    End If
     
End Sub
''Public Function Pan_(pname As String) As String
''   If rs1.State = 1 Then rs1.close
''   rs1.Open "select pan from SLEDGER where SUBLEDGER='" & pname & "'", con
''   If (Not IsNull(rs1(0)) Or rs1(0) <> "") Then
''      Pan_ = rs1(0)
''   Else
''      Pan_ = ""
''   End If
''End Function
Public Function MaxSNo_New(tbl As String, fld As String, c As String) As String
Dim RS As New Recordset

If RS.State = 1 Then RS.close
If c = "Country" Then
    
    RS.Open "Select max(" & fld & ") from " & tbl, con
    If IsNull(RS(0)) Then
        MaxSNo_New = "C0000" & 1
    Else
        MaxSNo_New = "C" & "" & Format(Val(Mid(RS(0), 2)) + 1, "00000")
    End If
    
ElseIf c = "Rep" Then
    
    RS.Open "Select max(" & fld & ") from " & tbl, con
    If IsNull(RS(0)) Then
        MaxSNo_New = "R0000" & 1
    Else
        MaxSNo_New = "R" & "" & Format(Val(Mid(RS(0), 2)) + 1, "00000")
    End If
    
    
ElseIf c = "state" Then
   
    RS.Open "Select max(" & fld & ") from " & tbl, con
    If IsNull(RS(0)) Then
        MaxSNo_New = "S0000" & 1
    Else
        MaxSNo_New = "S" & "" & Format(Val(Mid(RS(0), 2)) + 1, "00000")
    End If
   
ElseIf c = "District" Then

   RS.Open "Select max(" & fld & ") from " & tbl, con
    If IsNull(RS(0)) Then
        MaxSNo_New = "D0000" & 1
    Else
        MaxSNo_New = "D" & "" & Format(Val(Mid(RS(0), 2)) + 1, "00000")
    End If
 
ElseIf c = "city" Then

   RS.Open "Select max(" & fld & ") from " & tbl, con
    If IsNull(RS(0)) Then
        MaxSNo_New = "C0000" & 1
    Else
        MaxSNo_New = "C" & "" & Format(Val(Mid(RS(0), 2)) + 1, "00000")
    End If

ElseIf c = "University" Then


ElseIf c = "Department" Then

 
ElseIf c = "BookType" Then

 
ElseIf c = "Auther" Then

ElseIf c = "Product" Then

   RS.Open "Select max(" & fld & ") from " & tbl, con
   If IsNull(RS(0)) Then
        MaxSNo_New = "P0000" & 1
   Else
        MaxSNo_New = "P" & "" & Format(Val(Mid(RS(0), 2)) + 1, "00000")
   End If
   
ElseIf c = "college" Then

   RS.Open "Select max(" & fld & ") from " & tbl, con
   If IsNull(RS(0)) Then
        MaxSNo_New = "C0000" & 1
   Else
        MaxSNo_New = "C" & "" & Format(Val(Mid(RS(0), 2)) + 1, "00000")
   End If

ElseIf c = "collegeCode" Then

   RS.Open "Select  max(" & fld & ")   from " & tbl, con
   If IsNull(RS(0)) Then
        MaxSNo_New = "C-" & 1
   Else
        MaxSNo_New = "C-" & "" & Val(Mid(RS(0), 3)) + 1
   End If

ElseIf c = "teacher" Then
   
   RS.Open "Select max(" & fld & ") from " & tbl, con
   If IsNull(RS(0)) Then
        MaxSNo_New = "T0000" & 1
   Else
        MaxSNo_New = "T" & "" & Format(Val(Mid(RS(0), 2)) + 1, "00000")
   End If

ElseIf c = "teacherCode" Then

   RS.Open "Select  max(" & fld & ")   from " & tbl, con
   If IsNull(RS(0)) Then
        MaxSNo_New = "T-" & 1
   Else
        MaxSNo_New = "T-" & "" & Val(Mid(RS(0), 3)) + 1
   End If
   
'=========

End If


End Function
Public Function checkAuthentication(tbl As String, invhead As String, inv As Integer) As Boolean
    
checkAuthentication = False

End Function
Public Function Modifytbl(sledger As String, Newsledger As String)
      
con.Execute "update vouchers set SubLedger='" & Newsledger & "' where SubLedger='" & sledger & "' and " & stringyear
con.Execute "update invoicea set SubLedger='" & Newsledger & "' where SubLedger='" & sledger & "' and " & stringyear
con.Execute "update invoiceb set SubLedger='" & Newsledger & "' where SubLedger='" & sledger & "' and " & stringyear
con.Execute "update Casha set SubLedger='" & Newsledger & "' where SubLedger='" & sledger & "' and " & stringyear
con.Execute "update Cashb set SubLedger='" & Newsledger & "' where SubLedger='" & sledger & "' and " & stringyear
con.Execute "update CREDITA set SubLedger='" & Newsledger & "' where SubLedger='" & sledger & "' and " & stringyear
con.Execute "update CREDITB set SubLedger='" & Newsledger & "' where SubLedger='" & sledger & "' and " & stringyear

con.Execute "update DNFA set [PSLD]='" & Newsledger & "' where [PSLD]='" & sledger & "' and " & stringyear
con.Execute "update DNFB set [SLD]='" & Newsledger & "' where [SLD]='" & sledger & "' and " & stringyear

con.Execute "update CNF1A set [PSLD]='" & Newsledger & "' where [PSLD]='" & sledger & "' and " & stringyear
con.Execute "update CNF1B set [SLD]='" & Newsledger & "' where [SLD]='" & sledger & "' and " & stringyear
     
End Function
Public Function updateGledger(Newsledger As String, sledger As String)
      
      
con.Execute "update vouchers set GenLedger='" & Newsledger & "' where GenLedger='" & sledger & "' and " & stringyear
con.Execute "update invoicea set GENLEDGER='" & Newsledger & "' where GENLEDGER='" & sledger & "' and " & stringyear
con.Execute "update invoiceb set Genledger='" & Newsledger & "' where Genledger='" & sledger & "' and " & stringyear
con.Execute "update Casha set GENLEDGER='" & Newsledger & "' where GENLEDGER='" & sledger & "' and " & stringyear
con.Execute "update Cashb set Genledger='" & Newsledger & "' where Genledger='" & sledger & "' and " & stringyear
con.Execute "update CREDITA set GENLEDGER='" & Newsledger & "' where SubLedger='" & sledger & "' and " & stringyear
con.Execute "update CREDITB set GENLEDGER='" & Newsledger & "' where SubLedger='" & sledger & "' and " & stringyear

con.Execute "update DNFA set [PGLD]='" & Newsledger & "' where [PGLD]='" & sledger & "' and " & stringyear
con.Execute "update DNFB set [GLD]='" & Newsledger & "' where [GLD]='" & sledger & "' and " & stringyear



con.Execute "update CNF1A set [PGLD]='" & Newsledger & "' where [PGLD]='" & sledger & "' and " & stringyear
con.Execute "update CNF1B set [GLD]='" & Newsledger & "' where [GLD]='" & sledger & "' and " & stringyear



     
End Function
Function check_Duplikate(tbl As String, value_no As String) As Boolean

If RS.State = 1 Then RS.close

Select Case tbl


Case "invoicea"
        
      RS.Open "select top 1 * from " & tbl & " where invoiceno=" & value_no & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
      If RS.EOF = False Then
         check_Duplikate = True
      Else
         check_Duplikate = False
      End If
        

Case "credita"

      RS.Open "select * from " & tbl & " where invoiceno=" & value_no & "", con, adOpenKeyset, adLockReadOnly
      If RS.EOF = False Then
         check_Duplikate = True
      Else
         check_Duplikate = False
      End If


Case "casha"

      RS.Open "select * from " & tbl & " where invoiceno=" & value_no & "", con, adOpenKeyset, adLockReadOnly
      If RS.EOF = False Then
         check_Duplikate = True
      Else
         check_Duplikate = False
      End If



End Select

End Function
Sub UpdateDisPatchReg(inv As Long, invDt As String, pname As String, station As String, bundle As String, transport As String, marka As String, bilty As String, grData As String, frieght As String, tlb As String)
     
    Dim rs_dis As New ADODB.Recordset
    If rs_dis.State = 1 Then rs_dis.close
    rs_dis.Open "select * from " & tlb & " where sno=" & inv & "", con, adOpenDynamic, adLockOptimistic
    If rs_dis.EOF = True Then
       rs_dis.AddNew
       rs_dis!sno = inv
       rs_dis!Date = invDt
       rs_dis!Particulars = pname
       rs_dis!BDL = station
       rs_dis!wt = bundle
       rs_dis!freight = transport
       rs_dis!rr = marka
       rs_dis!gr = bilty
       If IsDate(grData) Then
       rs_dis!GR_DT = grData
       End If
       rs_dis!Freight_Paid = frieght
       rs_dis.update
    Else
        rs_dis!sno = inv
       rs_dis!Date = invDt
       rs_dis!Particulars = pname
       rs_dis!BDL = station
       rs_dis!wt = bundle
       rs_dis!freight = transport
       rs_dis!rr = marka
       rs_dis!gr = bilty
       If IsDate(grData) Then
       rs_dis!GR_DT = grData
       End If
       rs_dis!Freight_Paid = frieght
       rs_dis.update
    End If
     
End Sub

Function Party_Remove_FromOrder(party As String, gd As String, orderNo As Integer) As String
con.Execute "delete from OrderMnm where (name='" & party & "' and id=" & orderNo & ")"
End Function
Public Function ButtonPermission(cmdSave As CommandButton, cmdDelete As CommandButton, cmdedit As CommandButton) As Boolean
    
    
    
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select [Save],[Delete],[Edit] from UsrePermission where userName='" & main.UserName & "'", coninfo, adOpenKeyset, adLockReadOnly
    
    If rs1.EOF = False Then
       
       If rs1(0) = "y" Then
          cmdSave.Enabled = True
       Else
          cmdSave.Enabled = False
       End If
       
       If rs1(1) = "y" Then
          cmdDelete.Enabled = True
       Else
          cmdDelete.Enabled = False
       End If
       
       If rs1(2) = "y" Then
          cmdedit.Enabled = True
       Else
          cmdedit.Enabled = False
       End If
       
    
    End If
End Function
Sub BackColorFromNew(frm As Form, Optional c As String, Optional p As frame, Optional P1 As frame, Optional vs1 As VSFlexGrid)
    
    
   Dim o As Object
    
   Set rs1 = New ADODB.Recordset
   rs1.Open "select top 500 color_dark, color_light, Name " & _
   "FROM color_setting where Module='" & module_ & "' and apply = 1", con
   If rs1.EOF = False Then
   
        DoEvents
        frm.BackColor = rs1!color_light
        
        DoEvents
        DoEvents
   End If
End Sub

Sub BackColorFrom(frm As Form, Optional c As String, Optional p As frame, Optional P1 As frame, Optional vs1 As VSFlexGrid)
    
   On Error Resume Next
    
   Dim o As Object
    
   If module_ = "" Then
      module_ = "invoicing"
   End If
    
   Set rs1 = New ADODB.Recordset
   rs1.Open "select top 500 color_dark, color_light, Name " & _
   "FROM color_setting where Module='" & module_ & "' and apply = 1", con
   
   If rs1.EOF = False Then
   
   
        frm.BackColor = rs1!color_light
        
       
        
        DoEvents
        DoEvents
        
        frm.buttonFrame.BackColor = rs1!color_light
        frm.Picture5.BackColor = rs1!color_light
        
        
        
        For Each o In frm
            
            If (TypeOf o Is frame) Then
                o.BackColor = rs1!color_light
                o.Enabled = True
            End If
        
            If (TypeOf o Is Shape) Then
                o.BorderColor = rs1!color_dark
            End If
            
            If (TypeOf o Is Label) Then
                o.BackColor = rs1!color_dark
            End If
            
            If (TypeOf o Is MSFlexGrid) Then
                o.BackColorFixed = rs1!color_dark
            End If
            
            If (TypeOf o Is VSFlexGrid) Then
                o.BackColorFixed = rs1!color_dark
            End If
            
            If (TypeOf o Is MSHFlexGrid) Then
                o.BackColorFixed = rs1!color_dark
            End If
            
            
            
        
        Next
        
   
   Else
   
        frm.BackColor = &HB8E4F1
        frm.buttonFrame.BackColor = &HB8E4F1
        frm.Picture5.BackColor = &HB8E4F1
        
        For Each o In frm
            If (TypeOf o Is frame) Then
                o.BackColor = &HB8E4F1
                 o.Enabled = True
            End If
        Next
    
   End If
    
    
    
    For K = 0 To frm.vs.Cols
        vs1.Cell(flexcpFontSize, 0, K) = 13
    Next
    
    
'-----------------------------------------------------------------------------------------------
'----------------------------Resise Form--------------------------------------------------------
'If username = "admin" Then
'   c = "1"
'End If

If c = "1" Then
'Exit Sub

Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer
Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors

 MyForm.Height = 0
 MyForm.Width = 0

 ' Size of Form in Pixels at design resolution
 DesignX = 800
 DesignY = 600
 
 
 
 
 RePosForm = True   ' Flag for positioning Form
 DoResize = False   ' Flag for Resize Event
 ' Set up the screen values
 Xtwips = Screen.TwipsPerPixelX
 Ytwips = Screen.TwipsPerPixelY
 Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
 Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

 ' Determine scaling factors
 ScaleFactorX = (Xpixels / DesignX)
 ScaleFactorY = (Ypixels / DesignY)
 ScaleMode = 1  ' twips
 'Exit Sub  ' uncomment to see how Form1 looks without resizing
 Resize_For_Resolution ScaleFactorX, ScaleFactorY, frm
 Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
  "  by " + Str$(Ypixels)
  
 MyForm.Height = frm.Height  ' Remember the current size
 MyForm.Width = frm.Width
  
If Xpixels = 1024 And Ypixels = 768 Then

    leftAlign = 140
    leftAlign_cash = 150
    leftAlign_crnot_I = 500
    divBy = 3.5

ElseIf Xpixels = 1366 And Ypixels = 768 Then
    leftAlign = 230
    leftAlign_cash = 180
    leftAlign_crnot_I = 300
    
   divBy = 1.4
ElseIf Xpixels >= 1600 And Ypixels >= 900 Then
    leftAlign = 300
    leftAlign_crnot_I = 300
   divBy = 7.5
   
End If



    For K = 0 To frm.grid1.Cols
        frm.grid1.ColWidth(K) = (frm.grid1.ColWidth(K) + (frm.grid1.ColWidth(K) / divBy))
    Next
    
    For J = 0 To frm.DGrid.Cols
        frm.DGrid.ColWidth(K) = (frm.DGrid.ColWidth(K) + (frm.DGrid.ColWidth(K) / divBy))
    Next

    For K = 0 To frm.vs1.Cols
        frm.vs1.ColWidth(K) = (frm.vs1.ColWidth(K) + (frm.vs1.ColWidth(K) / divBy))
    Next



End If
    
End Sub

Sub clearFrom(frm As Form)
         
    On Error Resume Next
    
    Dim o As Object
    For Each o In frm
    
      If (TypeOf o Is textbox Or TypeOf o Is ComboBox) Then
         o.text = ""
      End If
    
    Next
    
   frm.cmdEdit_4.Enabled = False
   frm.cmdDelete_3.Enabled = False
   frm.cmdSave_2.Enabled = True
    
         
End Sub
Public Function MaxSNoNew(tbl As String, fld As String, cat As String) As Double
    Dim rs_ As New Recordset
    If rs_.State = 1 Then rs_.close
    rs_.Open "Select max(" & fld & ") from " & tbl & " where categories='" & cat & "'", con
    If IsNull(rs_(0)) Then
        MaxSNoNew = 1
    Else
        MaxSNoNew = Val(rs_(0)) + 1
    End If
    rs_.close
End Function

'Developer: Dinesh Saini
'Get Max + 1 Number from Perticular Number Field from a table

Public Function MaxSNo(tbl As String, fld As String) As Double
    Dim rs_ As New Recordset
    If rs_.State = 1 Then rs_.close
    rs_.Open "Select max(" & fld & ") from " & tbl, con
    If IsNull(rs_(0)) Then
        MaxSNo = 1
    Else
        MaxSNo = Val(rs_(0)) + 1
    End If
    rs_.close
End Function
Public Function MaxBookNo(n As Long) As String

Dim a, a1 As Integer
a = 0
a1 = 0

If n <= 50 Then
bookNo = Format(1, "00")
Else
a = n Mod 50
If a > 0 Then
a1 = n - a
a = a1 / 50
bookNo = a + 1
bookNo = Format(bookNo, "00")

Else
bookNo = n / 50
bookNo = Format(bookNo, "00")

End If

End If
    
    
End Function
Sub PartyWiseDis_Con()



Set con_LAST1 = New ADODB.Connection
If LCase(server_) = "server" Then
   con_LAST1.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & last_dbase & "; UID=" & sql_user & "; PWD=" & sql_pass
   con_LAST1.Open
End If



End Sub

Sub ConOpen()

On Error GoTo errtxt

Set CCON = New ADODB.Connection
Set coninfo = New ADODB.Connection
  
Dim apppath_ As String

'------------------------------------
Open App.Path + "\client.txt" For Input As #1
Line Input #1, apppath_
Close #1
  

If login.Combo1.text = "2015-16" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig.mdb"
ElseIf login.Combo1.text = "2016-17" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig_1617.MDB"
ElseIf login.Combo1.text = "2017-18" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig_1718.MDB"
ElseIf login.Combo1.text = "2018-19" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig_1819.MDB"
ElseIf login.Combo1.text = "2019-20" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig_1920.MDB"
ElseIf login.Combo1.text = "2020-21" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig_2021.MDB"
ElseIf login.Combo1.text = "2021-22" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig_2122.MDB"
ElseIf login.Combo1.text = "2022-23" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig_2223.MDB"
ElseIf login.Combo1.text = "2023-24" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig_2324.MDB"
ElseIf login.Combo1.text = "2024-25" Then
   CCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + apppath_ + "\dataconfig_2425.MDB"
    
End If

CCON.Open

coninfo.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database password=nniaj;Persist Security Info=false;Data Source=" + apppath_ + "\info.MDB"
coninfo.Open

 
errtxt:

If err.Number = 53 Then
    MsgBox "Please Copy Client.txt into Application Path..." & vbCrLf & err.DESCRIPTION, vbCritical
    End
End If

If err.Number <> 0 Then
    MsgBox "" & vbCrLf & err.DESCRIPTION, vbCritical
    End
End If
End Sub
Sub FatchClosing(Genledger As String, cust As String, date1 As String, date2 As String, totalraw As Integer)

Dim fyear1, setupid As String
Dim dr, CR, op As Double



If cust <> "" Then

con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , SLEDGER.YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "  FROM SLEDGER  where  gledger='" + Trim(Genledger) + "' and SLEDGER.SUBLEDGER = '" & cust & "' and sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid, p, adCmdText
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
& " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))  " _
& " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and genledger='" + Trim(Genledger) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  0 AS OPAMOUNTDEBIT,sum(Purchasea.NETAMOUNT) AS OPAMOUNTCREDIT  , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
& " FROM (SLEDGER LEFT JOIN purchasea ON (SLEDGER.SUBLEDGER = purchasea.SUBLEDGER) AND (SLEDGER.gledger = purchasea.GENLEDGER) AND (SLEDGER.fyear = purchasea.fyear) AND (SLEDGER.setupid = purchasea.setupid))  " _
& " where  sledger.fyear='" & main.session & "' and purchasea.setupid=" & main.setupid & " and  genledger='" + Trim(Genledger) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and convert(smalldatetime,INVOICEDATE,103) <convert(smalldatetime,'" + Trim(date1) + "',103) " & _
" GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
& " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid)) " _
& " where   sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & "  and gledger='" + Trim(Genledger) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
& " where   sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & "  and genledger='" + Trim(Genledger) + "' and SLEDGER.SUBLEDGER ='" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER) AND (SLEDGER.fyear = credita.fyear) AND (SLEDGER.setupid = credita.setupid)) " _
& " where sledger.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and genledger='" + Trim(Genledger) + "' and SLEDGER.SUBLEDGER='" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
& " WHERE sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and genledger='" + Trim(Genledger) + "' and SLEDGER.SUBLEDGER= '" & cust & "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
& " WHERE sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and genledger='" + Trim(Genledger) + "' and SLEDGER.SUBLEDGER= '" & cust & "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid) " _
& " WHERE sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(Genledger) + "'and CNF1A.PSLD = '" & cust & "'  and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid)" _
& " WHERE sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(Genledger) + "' and CNF1A.PSLD = '" & cust & "'  and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)  AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
& " WHERE sledger.fyear='" & main.session & "' and DNFA.setupid=" & main.setupid & " and DNFA.DC='D' and pgld = '" + Trim(Genledger) + "' and   DNFA.PSLD = '" & cust & "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
& " WHERE sledger.fyear='" & main.session & "' and DNFA.setupid=" & main.setupid & " and DNFA.DC='C' and pgld = '" + Trim(Genledger) + "'  and  DNFA.PSLD = '" & cust & "'   and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD)  AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid)" _
& " WHERE sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(Genledger) + "' and  CNF1B.SLD = '" & cust & "'    and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1) + "',103)  " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid) " _
& " WHERE sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC= 'C' and gld  = '" + Trim(Genledger) + "'  and  CNF1B.SLD = '" & cust & "'      and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1) + "',103)  " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD)  AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
& " WHERE sledger.fyear='" & main.session & "' and DNFB.setupid=" & main.setupid & " and DNFB.DC='D' and gld = '" + Trim(Genledger) + "'  and  DNFB.SLD = '" & cust & "'     And convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
& " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
& " WHERE sledger.fyear='" & main.session & "' and DNFB.setupid=" & main.setupid & " and DNFB.DC='C' and gld = '" + Trim(Genledger) + "'  and  DNFB.SLD = '" & cust & "'    and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1) + "',103) " _
& " GROUP BY SLEDGER.SUBLEDGER " _
& " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText

'-------------------------------- Opening Code Close ----------------------------------------------------------------------------
'-------------------------------- Opening Code Close ----------------------------------------------------------------------------
'-------------------------------- Opening Code Close ----------------------------------------------------------------------------

con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)   SELECT GenLedger, SubLedger, VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(Genledger) + "' and  Subledger = '" & cust & "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(Genledger) + "'  ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(Genledger) + "' and  Subledger = '" & cust & "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEA  where    " & stringyear & " and genledger ='" + Trim(Genledger) + "' and  Subledger = '" & cust & "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103) "
'change By Dinesh
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT Purchasea.GENLEDGER, Purchasea.SUBLEDGER, Purchasea.INVOICEDATE, 'I' AS Expr1, Purchasea.billNO, 'Purchase Invoice' , Purchasea.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM Purchasea  where    " & stringyear & " and genledger ='" + Trim(Genledger) + "' and  Subledger = '" & cust & "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103) "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(Genledger) + "' and  Subledger = '" & cust & "'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)   "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(Genledger) + "' and  Subledger = '" & cust & "'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)  AND CASHA.BAA <>0  "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM CREDITA  where    " & stringyear & " and genledger='" + Trim(Genledger) + "' and  Subledger = '" & cust & "'    and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103) "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & "   From DNFA  where    " & stringyear & " and Pgld ='" + Trim(Genledger) + "' and  Psld = '" & cust & "' and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid )   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1A where  " & stringyear & " and Pgld='" + Trim(Genledger) + "' and  Psld = '" & cust & "'    and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid ,fyear,setupid)   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & ",'" & main.session & "'," & main.setupid & " From DNFB  where  " & stringyear & " and gld='" + Trim(Genledger) + "' and  sld = '" & cust & "'   and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid ,fyear,setupid)   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1B where  " & stringyear & " and gld='" + Trim(Genledger) + "' and  sld = '" & cust & "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
con.Execute "insert into Treport ( Genledger,Subledger,openingbalance,userid,fyear,setupid ) SELECT '" + Trim(Genledger) + "' as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId,'" & main.session & "'," & main.setupid & " from subledgertrail where  " & stringyear & " GROUP BY SUBLEDGER;"

End If

d10 = d10 + 1



If d10 >= totalraw Then

dr = 0
CR = 0
op = 0

If RS.State = 1 Then RS.close
RS.Open "select subledger from [ExportData].[dbo].[treport] group by subledger", con, adOpenKeyset, adLockReadOnly
While RS.EOF = False

CR = 0
dr = 0
sum1 = 0
op = 0

If rs1.State = 1 Then rs1.close
rs1.Open "select subledger,OpeningBalance,ad,dorc from [ExportData].[dbo].[treport] where subledger='" & RS(0) & "'", con, adOpenKeyset, adLockReadOnly
While rs1.EOF = False
    If rs1!dorc = "D" Then
      dr = dr + rs1!ad
    End If
    If rs1!dorc = "C" Then
      CR = CR + rs1!ad
    End If
    If rs1!OpeningBalance <> 0 Then
    op = rs1!OpeningBalance
    End If
rs1.MoveNext
Wend
sum1 = op + (dr - CR)

con.Execute "update sledger set balance = " & sum1 & "  Where  " & stringyear & " and SUBLEDGER='" & RS!subledger & "'"
RS.MoveNext
Wend


End If

End Sub
Function checkdate(datestr As String, ctl As Control) As Boolean

Dim check As Boolean
check = True
Dim kk As Integer
If IsDate(ctl) Then
    
    Dim RS As Recordset
    Set RS = New ADODB.Recordset
    RS.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockPessimistic, adCmdText
    If CDate(ctl) >= RS!yarfrom And CDate(ctl) <= RS!yarto Then
    Else
    check = False
    MsgBox "Enter Right Date! Considering This Session !!", vbInformation
    
    'End If
      
    End If
    RS.close
    For kk = 1 To Trim(Len(datestr))
        If Mid$(Trim(datestr), kk, 1) = "_" Then
            check = False
        End If
    Next
    If Not check Then
        checkdate = False
        Exit Function
    End If
cagain:
    If Len(Trim(datestr)) > 10 Then
        checkdate = False
    Else
        If Len(Trim(datestr)) < 10 Then
            If Mid$(Trim(datestr), 3, 1) = "/" Then
                If Mid$(Trim(datestr), 6, 1) = "/" Then
                    If Val(Mid$(Trim(datestr), 7, 2)) >= 85 And Val(Mid$(Trim(datestr), 7, 2)) <= 99 Then
                        kk = Val(Mid$(Trim(datestr), 7, 2))
                        datestr = Trim(Mid$(Trim(datestr), 1, 6)) + Trim("19")
                        If Len(Trim(Str(kk))) < 2 Then
                            datestr = Trim(datestr) + Trim("0") + Trim(Str(kk))
                        Else
                            datestr = Trim(datestr) + Trim(Str(kk))
                        End If
                        ctl.text = datestr
                    Else
                        kk = Val(Mid$(Trim(datestr), 7, 2))
                        datestr = Trim(Mid$(Trim(datestr), 1, 6)) + Trim("20")
                        If Len(Trim(Str(kk))) < 2 Then
                            datestr = Trim(datestr) + Trim("0") + Trim(Str(kk))
                        Else
                            datestr = Trim(datestr) + Trim(Str(kk))
                        End If
                        ctl.text = datestr
                    End If
                    GoTo cagain
                Else
                    checkdate = False
                End If
            Else
                checkdate = False
            End If
        Else
            If Val(Mid$(Trim(datestr), 1, 2)) > 31 Then
                        checkdate = False
            Else
                If Mid$(Trim(datestr), 3, 1) <> "/" Then
                    checkdate = False
                Else
                    If Mid$(Trim(datestr), 6, 1) <> "/" Then
                        checkdate = False
                    Else
                        If Val(Mid$(Trim(datestr), 4, 2)) > 12 Then
                            checkdate = False
                        Else
                            checkdate = True
                        End If
                    End If
                End If
            End If
        End If
    End If
Else
checkdate = False
End If
End Function
Function dspace(s As String) As String
Dim L, I As Integer
s = Trim(s)
dspace = ""
For I = 1 To Len(Trim(s))
    dspace = dspace + Mid$(s, I, 1) + " "
Next
End Function
Function cnullstr(s As String) As String
    If s = Null Then
        cnullstr = " "
    Else
        cnullstr = Trim(s)
    End If
End Function

Public Function printnow()
    Dim X As Long
    Dim p As Printer
    For I = 0 To Printers.Count - 1
        Set p = Printers(I)
        If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.text) Then
            Exit For
        End If
    Next
    For I = 1 To (Printdlg.UpDown1.value)
        X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(p.Port), vbHide)
    Next
    Printdlg.UpDown1.value = 1
    Printdlg.Text1.text = "1"
End Function
Public Function reportdata()
   'Set repocon = New ADODB.Connection
    Set repors = New ADODB.Recordset
    repors.Open "select * from printsetup ", con, adOpenStatic, adLockOptimistic, adCmdText
End Function
Public Function LEFTM() As Integer
    LEFTM = (main.repors!LeftMargin)
End Function
Public Function RIGHTM() As Integer
    RIGHTM = main.repors!RightMargin
End Function
Public Sub CNSetup()
   Dim crs As ADODB.Recordset
   Set crs = New ADODB.Recordset
   crs.Open "Select * from setup1 where " & stringyear & "", con, adOpenStatic, adLockOptimistic
   Dim RNO As Integer
   RNO = setupid  '4 + setupid
   Select Case RNO
       Case 2
            crs!cname = "Blueprint Education"
            crs!add1 = "(A division of Chitra Prakashan (I) Pvt.Ltd.)"
            crs!add2 = "513, Mohkampur Indl. Area, Phase II"
            crs.update
            fromDate_setup = crs!yarfrom
            toDate_setup = crs!yarto
            
            billformat = crs!billformat & ""
            itmeCode = crs!Code & ""
            gst = crs!cst
            
            firm_Address = "Blueprint Education (A division of Chitra Prakashan (I) Pvt.Ltd.)  513, Mohkampur Indl. Area, Phase II"
            
          
            
       Case 1
            crs!cname = "CHITRA PRAKASHAN (I) PVT.LTD."
            crs!add1 = "Western Kutchery Road,Meerut"
            crs.update
            
            fromDate_setup = crs!yarfrom
            toDate_setup = crs!yarto
            
            
       Case 3
            crs!cname = "DISHA SURGICALS (P) LTD."
            crs!add1 = "Manufactures, Suppliers & Exporters of Surgical Dressings"
            crs!add2 = "145, SAI PURAM, DELHI ROAD, MEERUT"
            crs.update
       Case 5
            crs!cname = "Zenith Corportion"
            App.Title = crs!cname
            crs!add1 = "5-A Jain Nagar, Meerut -250002 (U.P.)"
            crs!add2 = ""
            crs.update
       
       Case 6
            crs!cname = "Zest Surgical Pvt. Ltd."
            App.Title = crs!cname
            crs!add1 = "47-B Jain Nagar, Meerut -250002 (U.P.)"
            crs!add2 = ""
            crs.update
       Case 7
            crs!cname = "Indu Surgical Mfg. Co."
            App.Title = crs!cname
            crs!add1 = "34-B Jain Nagar, Meerut -250002 (U.P.)"
            crs!add2 = ""
            crs.update
       Case 8
            crs!cname = "ZestKK"
            App.Title = crs!cname
            crs!add1 = "47-B Jain Nagar, Meerut -250002 (U.P.)"
            crs!add2 = ""
            crs.update
   End Select
End Sub
Sub popuplistModel10(ByVal ST As String, ByRef cn1 As ADODB.Connection, Optional ar As Integer, Optional COLLAGE As Boolean)

Set rs1 = New ADODB.Recordset
rs1.Open ST, cn1, adOpenForwardOnly, adLockReadOnly

If rs1.EOF = False Then
        
        popuplistModel.lblRaw.Caption = rs1.RecordCount
        
        If ar = 0 Then
            ar = rs1.Fields.Count
        ElseIf ar > rs1.Fields.Count Then
            ar = rs1.Fields.Count
        End If
        
        ReDim Array1(ar) As Variant
        For q = 0 To ar - 1
            Array1(q) = 0
        Next q
        Unload popuplistModel
        
        popuplistModel.ListView1.ListItems.Clear
        
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        popupvalue4 = ""
            popuplistModel.lblRaw.Caption = " Total Row : " & rs1.RecordCount
            For I = 1 To ar
            popuplistModel.ListView1.ColumnHeaders.Add I, , rs1.Fields(I - 1).Name
            popuplistModel.ListView1.ColumnHeaders(I).Width = 2000
            
            
            
            
            
        If searchType = "paper" Then
            If I = 1 Then
               popuplistModel.ListView1.ColumnHeaders(I).Width = 1200
            ElseIf I = 2 Then
               popuplistModel.ListView1.ColumnHeaders(I).Width = 1200
            ElseIf I = 3 Then
               popuplistModel.ListView1.ColumnHeaders(I).Width = 1200
               
            End If
        
            
        Else
            If I = 1 Then
               popuplistModel.ListView1.ColumnHeaders(I).Width = 3500
            ElseIf I = 2 Then
               popuplistModel.ListView1.ColumnHeaders(I).Width = 3000
            End If
            
        End If
            
            
        Next I
        
        
        popuplistModel.ListView1.View = lvwReport
        Dim LItem As ListItem
        If COLLAGE = True Then
            Progress.Show
            Progress.pb1 = 0
            Progress.pb1.Max = rs1.RecordCount
        End If
        While Not rs1.EOF
        If Not IsNull(rs1.Fields(0)) Then
            Set LItem = popuplistModel.ListView1.ListItems.Add(, , rs1.Fields(0).value)
            If Len(rs1.Fields(0).value) > Len(rs1.Fields(0).Name) Then
               Array1(0) = Len(rs1.Fields(0).value)
            Else
               Array1(0) = Len(rs1.Fields(0).Name)
            End If
            If ar > 0 Then
                For m = 1 To ar - 1
                   If Not IsNull(rs1.Fields(m).value) Then LItem.SubItems(m) = rs1.Fields(m).value
                    If Len(rs1.Fields(m).value) > Len(rs1.Fields(m).Name) Then
                       Array1(m) = Len(rs1.Fields(m).value)
                    Else
                       Array1(m) = Len(rs1.Fields(m).Name)
                    End If
                Next m
            End If
            End If
            rs1.MoveNext
            If COLLAGE = True Then
                Progress.pb1 = Progress.pb1 + 1
                If Progress.pb1 = Progress.pb1.Max Then Unload Progress
            End If
        Wend
     '   For z = 1 To ar
     '        If Array1(z - 1) <= 2 Then Array1(z - 1) = 4
     '        popuplistModel.ListView1.ColumnHeaders(z).Width = Array1(z - 1) * 150
     '   Next z
     
     popuplistModel.Show 1
     
End If
rs1.close
End Sub
Sub popuplist_SP(ByVal ST As String, ByRef cn1 As ADODB.Connection, Optional ar As Integer, Optional COLLAGE As Boolean)

Set rs1 = New ADODB.Recordset

Set rs1 = con.Execute("" & ST)

'rs1.Open ST, cn1, adOpenForwardOnly, adLockReadOnly

If rs1.EOF = False Then
        popuplistModel.lblRaw.Caption = rs1.RecordCount
        If ar = 0 Then
            ar = rs1.Fields.Count
        ElseIf ar > rs1.Fields.Count Then
            ar = rs1.Fields.Count
        End If
        ReDim Array1(ar) As Variant
        For q = 0 To ar - 1
            Array1(q) = 0
        Next q
        ''Unload popuplist
        
        popuplistModel.ListView1.ListItems.Clear
        
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        popupvalue4 = ""
            popuplistModel.lblRaw.Caption = " Total Row : " & rs1.RecordCount
            For I = 1 To ar
            popuplistModel.ListView1.ColumnHeaders.Add I, , rs1.Fields(I - 1).Name
            popuplistModel.ListView1.ColumnHeaders(I).Width = 2000
            
            
            
            
            
        If searchType = "inv" Then
        
            
        Else
            If I = 1 Then
               popuplistModel.ListView1.ColumnHeaders(I).Width = 3500
            ElseIf I = 2 Then
               popuplistModel.ListView1.ColumnHeaders(I).Width = 3000
            End If
            
        End If
            
            
        Next I
        
        
        popuplistModel.ListView1.View = lvwReport
        Dim LItem As ListItem
        If COLLAGE = True Then
            Progress.Show
            Progress.pb1 = 0
            Progress.pb1.Max = rs1.RecordCount
        End If
        While Not rs1.EOF
        If Not IsNull(rs1.Fields(0)) Then
            Set LItem = popuplistModel.ListView1.ListItems.Add(, , rs1.Fields(0).value)
            If Len(rs1.Fields(0).value) > Len(rs1.Fields(0).Name) Then
               Array1(0) = Len(rs1.Fields(0).value)
            Else
               Array1(0) = Len(rs1.Fields(0).Name)
            End If
            If ar > 0 Then
                For m = 1 To ar - 1
                   If Not IsNull(rs1.Fields(m).value) Then LItem.SubItems(m) = rs1.Fields(m).value
                    If Len(rs1.Fields(m).value) > Len(rs1.Fields(m).Name) Then
                       Array1(m) = Len(rs1.Fields(m).value)
                    Else
                       Array1(m) = Len(rs1.Fields(m).Name)
                    End If
                Next m
            End If
            End If
            rs1.MoveNext
            If COLLAGE = True Then
                Progress.pb1 = Progress.pb1 + 1
                If Progress.pb1 = Progress.pb1.Max Then Unload Progress
            End If
        Wend
        
     '   For z = 1 To ar
     '        If Array1(z - 1) <= 2 Then Array1(z - 1) = 4
     '        popuplistModel.ListView1.ColumnHeaders(z).Width = Array1(z - 1) * 150
     '   Next z
     popuplistModel.Show 1
End If
rs1.close
End Sub

Sub popuplist10(ByVal ST As String, ByRef cn1 As ADODB.Connection, Optional ar As Integer, Optional COLLAGE As Boolean, Optional type_ As String)
Set rs1 = New ADODB.Recordset

rs1.Open ST, cn1, adOpenForwardOnly, adLockReadOnly

If rs1.EOF = False Then
        popuplist.lblRaw.Caption = rs1.RecordCount
        If ar = 0 Then
            ar = rs1.Fields.Count
        ElseIf ar > rs1.Fields.Count Then
            ar = rs1.Fields.Count
        End If
        ReDim Array1(ar) As Variant
        For q = 0 To ar - 1
            Array1(q) = 0
        Next q
        Unload popuplist
        
        popuplist.ListView1.ListItems.Clear
        
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        popupvalue4 = ""
            popuplist.lblRaw.Caption = " Total Row : " & rs1.RecordCount
            For I = 1 To ar
            popuplist.ListView1.ColumnHeaders.Add I, , rs1.Fields(I - 1).Name
            
            
            If searchType = "inv" Then
                bill = "0,1000,1200,5500,2000"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)
            ElseIf searchType = "party" Then
                bill = "0,3500,3200,3200,2000,2000"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)
            ElseIf searchType = "books" Then
                bill = "0,1500,6500,100,100,100"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)

            Else
               popuplist.ListView1.ColumnHeaders(I).Width = 3800
            End If
            
            
       
            
            
        Next I
        
        
        popuplist.ListView1.View = lvwReport
        Dim LItem As ListItem
        If COLLAGE = True Then
            Progress.Show
            Progress.pb1 = 0
            Progress.pb1.Max = rs1.RecordCount
        End If
        While Not rs1.EOF
        If Not IsNull(rs1.Fields(0)) Then
            Set LItem = popuplist.ListView1.ListItems.Add(, , rs1.Fields(0).value)
            If Len(rs1.Fields(0).value) > Len(rs1.Fields(0).Name) Then
               Array1(0) = Len(rs1.Fields(0).value)
            Else
               Array1(0) = Len(rs1.Fields(0).Name)
            End If
            If ar > 0 Then
                For m = 1 To ar - 1
                   If Not IsNull(rs1.Fields(m).value) Then LItem.SubItems(m) = rs1.Fields(m).value
                    If Len(rs1.Fields(m).value) > Len(rs1.Fields(m).Name) Then
                       Array1(m) = Len(rs1.Fields(m).value)
                    Else
                       Array1(m) = Len(rs1.Fields(m).Name)
                    End If
                Next m
            End If
            End If
            rs1.MoveNext
            If COLLAGE = True Then
                Progress.pb1 = Progress.pb1 + 1
                If Progress.pb1 = Progress.pb1.Max Then Unload Progress
            End If
        Wend
       
       popuplist.Show 1
       
End If

If rs1.State = 1 Then rs1.close
End Sub
Sub popuplistFast(ByVal ST As String, ByRef cn1 As ADODB.Connection, Optional ar As Integer, Optional COLLAGE As Boolean, Optional type_ As String)

Set rs1 = New ADODB.Recordset
Set rs1 = con.Execute("exec searchList '" & type_ & "'")

If rs1.EOF = False Then
        popuplist.lblRaw.Caption = rs1.RecordCount
        If ar = 0 Then
            ar = rs1.Fields.Count
        ElseIf ar > rs1.Fields.Count Then
            ar = rs1.Fields.Count
        End If
        ReDim Array1(ar) As Variant
        For q = 0 To ar - 1
            Array1(q) = 0
        Next q
        Unload popuplist
        
        popuplist.ListView1.ListItems.Clear
        
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        popupvalue4 = ""
            popuplist.lblRaw.Caption = " Total Row : " & rs1.RecordCount
            For I = 1 To ar
            popuplist.ListView1.ColumnHeaders.Add I, , rs1.Fields(I - 1).Name
            
            
            If searchType = "inv" Then
                bill = "0,1000,1200,5500,2000"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)
            ElseIf searchType = "party" Then
                bill = "0,3500,3200,3200,2000,2000"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)
            ElseIf searchType = "bookretailer" Then
                bill = "3500,3200,3200,2000,2000"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)
            
            ElseIf searchType = "books" Then
                bill = "0,1500,6500,100,100,100"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)
            ElseIf searchType = "cmaster" Then
                bill = "6000,6000,6500,600,600,600"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)
            ElseIf searchType = "inv1" Then
                bill = "0,1000,4000,1500,3500,1200"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)
            ElseIf searchType = "inv11" Then
                bill = "0,4000,1000,1500,3500,1200"
                a_strResult = Split(bill, ",")
                popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)

            Else
               popuplist.ListView1.ColumnHeaders(I).Width = 3000
            End If
            
            
       
            
            
        Next I
        
        
        popuplist.ListView1.View = lvwReport
        Dim LItem As ListItem
        If COLLAGE = True Then
            Progress.Show
            Progress.pb1 = 0
            Progress.pb1.Max = rs1.RecordCount
        End If
        While Not rs1.EOF
        If Not IsNull(rs1.Fields(0)) Then
            Set LItem = popuplist.ListView1.ListItems.Add(, , rs1.Fields(0).value)
            If Len(rs1.Fields(0).value) > Len(rs1.Fields(0).Name) Then
               Array1(0) = Len(rs1.Fields(0).value)
            Else
               Array1(0) = Len(rs1.Fields(0).Name)
            End If
            If ar > 0 Then
                For m = 1 To ar - 1
                   If Not IsNull(rs1.Fields(m).value) Then LItem.SubItems(m) = rs1.Fields(m).value
                    If Len(rs1.Fields(m).value) > Len(rs1.Fields(m).Name) Then
                       Array1(m) = Len(rs1.Fields(m).value)
                    Else
                       Array1(m) = Len(rs1.Fields(m).Name)
                    End If
                Next m
            End If
            End If
            rs1.MoveNext
            If COLLAGE = True Then
                Progress.pb1 = Progress.pb1 + 1
                If Progress.pb1 = Progress.pb1.Max Then Unload Progress
            End If
        Wend
       
       popuplist.Show 1
       
End If
rs1.close
End Sub
Sub popuplistFastNew(ByVal ST As String, ByRef cn1 As ADODB.Connection, Optional ar As Integer, Optional COLLAGE As Boolean, Optional type_ As String)

Set rs1 = New ADODB.Recordset
Set rs1 = con.Execute("exec Sp_FetchSaleReturn '" & last_dbase & "',''")

If rs1.EOF = False Then
        popuplist.lblRaw.Caption = rs1.RecordCount
        If ar = 0 Then
            ar = rs1.Fields.Count
        ElseIf ar > rs1.Fields.Count Then
            ar = rs1.Fields.Count
        End If
        ReDim Array1(ar) As Variant
        For q = 0 To ar - 1
            Array1(q) = 0
        Next q
        Unload popuplist
        
        popuplist.ListView1.ListItems.Clear
        
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        popupvalue4 = ""
            popuplist.lblRaw.Caption = " Total Row : " & rs1.RecordCount
            For I = 1 To ar
            popuplist.ListView1.ColumnHeaders.Add I, , rs1.Fields(I - 1).Name
            
      
            bill = "0,1200,1200,1200,1500,4000"
            a_strResult = Split(bill, ",")
            popuplist.ListView1.ColumnHeaders(I).Width = a_strResult(I)
            ''popuplist.ListView1.ColumnHeaders(I).Width = 1500
        
            
            
       
            
            
        Next I
        
        
        popuplist.ListView1.View = lvwReport
        Dim LItem As ListItem
        If COLLAGE = True Then
            Progress.Show
            Progress.pb1 = 0
            Progress.pb1.Max = rs1.RecordCount
        End If
        While Not rs1.EOF
        If Not IsNull(rs1.Fields(0)) Then
            Set LItem = popuplist.ListView1.ListItems.Add(, , rs1.Fields(0).value)
            If Len(rs1.Fields(0).value) > Len(rs1.Fields(0).Name) Then
               Array1(0) = Len(rs1.Fields(0).value)
            Else
               Array1(0) = Len(rs1.Fields(0).Name)
            End If
            If ar > 0 Then
                For m = 1 To ar - 1
                   If Not IsNull(rs1.Fields(m).value) Then LItem.SubItems(m) = rs1.Fields(m).value
                    If Len(rs1.Fields(m).value) > Len(rs1.Fields(m).Name) Then
                       Array1(m) = Len(rs1.Fields(m).value)
                    Else
                       Array1(m) = Len(rs1.Fields(m).Name)
                    End If
                Next m
            End If
            End If
            rs1.MoveNext
            If COLLAGE = True Then
                Progress.pb1 = Progress.pb1 + 1
                If Progress.pb1 = Progress.pb1.Max Then Unload Progress
            End If
        Wend
       
       popuplist.Show 1
       
End If
rs1.close
End Sub

Function myround(ByVal roundval As String, ByVal decplace As Byte) As Double
myround = Round(Val(roundval) + 0.000005, decplace)
End Function
Sub uploadData(dates As Date)



Dim kk As Integer
Dim ss As String



con.Execute "Delete from StockRegister where len(bookcode)>0"


extra_godown = "Za"

'''''''''''''''''' Sales------------------------------------------------------------------------------------
'convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date1.Text) + "',103)

'''CON.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) " & _
'''" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'Sales','Issue',inva.Godown  from INVOICEB invb " & _
'''" inner join invoicea as inva on  invb.invoiceno = inva.invoiceno " & _
'''" where convert(smalldatetime,invb.INVOICEDATE,103)>=onvert(smalldatetime,'" + Dates + "',103) and inva.Godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.Godown"


con_sales.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'SalesReturn','Receive',inva.Godown  from CREDITB invb " & _
" inner join CREDITa as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.Godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.Godown"

con_sales.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'CashSale','Issue',inva.Godown  from cashb invb " & _
" inner join casha as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.Godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.Godown"

'''''''''''''''''' Sales II------------------------------------------------------------------------------------

con_sales2.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "' " & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'Sales','Issue',inva.Godown  from INVOICEB invb " & _
" inner join invoicea as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.Godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.Godown"

con_sales2.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'SalesReturn','Receive',inva.Godown  from CREDITB invb " & _
" inner join CREDITa as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.Godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.Godown"

con_sales2.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'CashSale','Issue',inva.Godown  from cashb invb " & _
" inner join casha as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.Godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.Godown"


'**************************************************************************************************************
'--------------------------------- Code For Specimen----------------------------------------------------------


con_conven.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "' " & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'Specimen','Issue',inva.godown  from INVOICEB invb " & _
" inner join invoicea as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.godown"

con_conven.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'SpecimenReturn','Receive',inva.godown  from CREDITB invb " & _
" inner join CREDITa as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.godown"

'--------------------------------- Code For Specimen II----------------------------------------------------------

con_spII.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "' " & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'Specimen','Issue',inva.godown  from INVOICEB invb " & _
" inner join invoicea as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.godown"

con_spII.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'SpecimenReturn','Receive',inva.godown  from CREDITB invb " & _
" inner join CREDITa as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.godown<>'" & extra_godown & "'  group by  invb.INVOICEDATE, invb.BookCode,inva.godown"

'--------------------------------- Code For Basil Software-----------------------------------------------------


con_basil.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "' " & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'Sales','Issue',inva.Godown  from cashB invb " & _
" inner join casha as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.Godown<>'" & extra_godown & "' group by  invb.INVOICEDATE, invb.BookCode,inva.Godown"

con_basil.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select  invb.INVOICEDATE, invb.BookCode,sum(Quantity),'SalesReturn','Receive',inva.Godown  from Ret_CASHB invb " & _
" inner join Ret_CASHA as inva on  invb.invoiceno = inva.invoiceno where invb.INVOICEDATE>=datevalue('" & dates & "') and inva.Godown<>'" & extra_godown & "' group by  invb.INVOICEDATE, invb.BookCode,inva.Godown"


'--------- basil closing
Dim d1
d1 = Format(dates, "dd/MM/yyyy")

con_basil.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "' " & _
" select  '" & d1 & "', BookCode,qty,'SalesReturn','Receive',Godown  from bookclosing where Qty>0"

con_basil.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "' " & _
" select  '" & d1 & "', BookCode,(-1*qty),'Sales','Issue',Godown  from bookclosing where Qty<0"

'--------------------------------- Code For Binder Software-----------------------------------------------------

con_Binder.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom,BinderName) IN '" & ss & "'" & _
" select inv.INVOICEDATE,inv.book_code,sum(inv.NetBook),'Binder','Issue',inva.godown,inva.SUBLEDGER from invoiceb as inv inner join invoicea as inva " & _
" on inv.invoiceno = inva.INVOICENO where inv.INVOICEDATE>=datevalue('" & dates & "') and inva.godown<>'" & extra_godown & "' group by inv.INVOICEDATE,inv.Book_Code,inva.SUBLEDGER,inva.godown "

'Book Rec

con.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom,BinderName) IN '" & ss & "'" & _
" select inv.INVOICEDATE,inv.book_code,sum(inv.NetBook),'Binder','Receive',inva.godown,inva.SUBLEDGER from BookReceiveDet as inv inner join BinderBkReceive as inva " & _
" on inv.invoiceno = inva.INVOICENO where inv.INVOICEDATE>=datevalue('" & dates & "') and inva.godown<>'" & extra_godown & "' group by inv.INVOICEDATE,inv.Book_Code,inva.SUBLEDGER,inva.godown "



'--------------------------------------------------------------------------------------------------------------

con.Execute "delete from BinderMaster where len(binder_id)>0"

If RS.State = 1 Then RS.close
RS.Open "select SUBLEDGER,ADDRESS1,ADDRESS2,Contact_P,Phone,Mobile from SLEDGER", con_Binder
While RS.EOF = False
   kk = MaxSNo("BinderMaster", "Binder_id")
   con.Execute "insert into BinderMaster(Binder_id,Binder_name,add1) values(" & kk & ",'" & RS!subledger & "','" & RS!address1 & "')"
   RS.MoveNext
Wend


'''--------------------------------- End code  ------------------------------------------------------------------
'****************************************************************************************************************

con.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom,BinderName) IN '" & ss & "'" & _
" select Dates,BookCode,sum(Qty),Category,Issue_Receive,iif(Godown_In='-',Godown_out,Godown_in),Binder_Code from BookStock where dates>=datevalue('" & dates & "') and Godown_In='-' and len(Binder_Code)>0 and Godown_Out<>'" & extra_godown & "' group by DATEs,BookCode,Category,Issue_Receive,Binder_Code,Godown_Out,Godown_in"

con.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom,BinderName) IN '" & ss & "'" & _
" select Dates,BookCode,sum(Qty),Category,Issue_Receive,iif(Godown_In='-',Godown_out,Godown_in),Binder_Code from BookStock where dates>=datevalue('" & dates & "') and Godown_Out='-' and len(Binder_Code)>0 and Godown_in<>'" & extra_godown & "' group by DATEs,BookCode,Category,Issue_Receive,Binder_Code,Godown_Out,Godown_in"

con.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select Dates,BookCode,sum(Qty),Category,Issue_Receive,Godown_Out from BookStock where dates>=datevalue('" & dates & "') and Godown_In='-' and len(Binder_Code)=0 and Godown_Out<>'" & extra_godown & "' group by DATEs,BookCode,Category,Issue_Receive,Binder_Code,Godown_Out"


''
con.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select Dates,BookCode,sum(Qty),Category,Issue_Receive,Godown_Out from BookStock where dates>=datevalue('" & dates & "') and Godown_In<>'-' and Godown_out<>'-' and len(Binder_Code)=0 and Godown_Out<>'" & extra_godown & "' group by DATEs,BookCode,Category,Issue_Receive,Binder_Code,Godown_Out"

con.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
" select Dates,BookCode,sum(Qty),Category,'Receive',Godown_in from BookStock where dates>=datevalue('" & dates & "') and Godown_In<>'-' and Godown_out<>'-' and len(Binder_Code)=0 and Godown_Out<>'" & extra_godown & "' group by DATEs,BookCode,Category,Issue_Receive,Binder_Code,Godown_in"



'CON.Execute "insert into  StockRegister(Dates,BookCode,Qty,Category,Issue_Receive,Issue_ReceveFrom) IN '" & ss & "'" & _
'" select Dates,BookCode,sum(Qty),Category,Issue_Receive,Godown_Out from BookStock where Godown_In<>'-' and len(Binder_Code)=0 and Godown_in<>'" & extra_godown & "' group by DATEs,BookCode,Category,Issue_Receive,Binder_Code,Godown_Out"



'Exit Sub
'err_connection:
'MsgBox "" & Err.Description
     
     
End Sub
Sub popuplist1(ByVal ST As String, ByRef cn1 As ADODB.Connection, Optional ar As Integer, Optional font2 As String)
On Error Resume Next
 
Set rs1 = New ADODB.Recordset

If ar = 0 Then
rs1.Open ST, con, adOpenStatic, adLockReadOnly
Else
rs1.Open ST, CCON, adOpenStatic, adLockReadOnly
End If


If font2 = "h" Then
    Dim fill As New ADODB.Recordset
    Set fill = New ADODB.Recordset
    
    'popuplistVS.vs.FontSize = 12
    
    fill.Open ST, cn1
    
    Set popuplistVS.vs.DataSource = fill
    'popuplistvs.vs.Col = 2
    '-------------------------------
    
        For I = 0 To popuplistVS.vs.rows - 1
               DoEvents
            If font2 = "h" Then
               DoEvents
               DoEvents
               popuplistVS.vs.Cell(flexcpFontName, I, 1) = "Kundli"
               popuplistVS.vs.Cell(flexcpFontSize, I, 1) = 15
               
            End If
        Next
        
    '----------------------------------
    
     popuplistVS.vs.ColWidth(0) = 1500
     popuplistVS.vs.ColWidth(1) = 2800
     popuplistVS.vs.ColWidth(2) = 1600
     popuplistVS.vs.ColWidth(3) = 1600
     popuplistVS.vs.ColWidth(4) = 1600

    
    popuplistVS.Show
    Exit Sub
End If
If rs1.RecordCount > 0 Then
         
        If ar = 0 Then
            ar = rs1.Fields.Count
        ElseIf ar > rs1.Fields.Count Then
            ar = rs1.Fields.Count
        End If
        ReDim Array1(ar) As Variant
        For q = 0 To ar - 1
            Array1(q) = 0
        Next q
        Unload popuplistModel
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        
            For I = 1 To ar
            popuplistModel.ListView1.ColumnHeaders.Add I, , Trim(rs1.Fields(I - 1).Name)
            If I = 1 Then
               popuplistModel.ListView1.ColumnHeaders(I).Width = 2400
            Else
               popuplistModel.ListView1.ColumnHeaders(I).Width = 2600
            End If
            
           Select Case GridWidthSet
            
           Case 1:
                If I = 1 Then
                   popuplistModel.ListView1.ColumnHeaders(I).Width = 1500
                End If
           Case 2:
           
           Case 3:
           
           Case 4:
            
           End Select
        
        Next I
        
        GridWidthSet = 0
        
        
        
         If font2 = "h" Then
        'If frmbook.bfont = "h" Then
            popuplistModel.ListView1.Font = hindi
            popuplistModel.ListView1.Font.Size = 15

        Else
            popuplistModel.ListView1.Font = english
            popuplistModel.ListView1.Font.Size = 12
        End If

        
        
        
        popuplistModel.ListView1.View = lvwReport
        Dim LItem As ListItem
        While Not rs1.EOF
        If Not IsNull(rs1.Fields(0)) Then
            Set LItem = popuplistModel.ListView1.ListItems.Add(, , rs1.Fields(0).value)
            
            If Len(rs1.Fields(0).value) > Len(rs1.Fields(0).Name) Then
               Array1(0) = Len(rs1.Fields(0).value)
            Else
               Array1(0) = Len(rs1.Fields(0).Name)
            End If
            If ar > 0 Then
                For m = 1 To ar - 1
                   If Not IsNull(rs1.Fields(m).value) Then LItem.SubItems(m) = rs1.Fields(m).value
                    If Len(rs1.Fields(m).value) > Len(rs1.Fields(m).Name) Then
                       Array1(m) = Len(rs1.Fields(m).value)
                    Else
                       Array1(m) = Len(rs1.Fields(m).Name)
                    End If
                    
                  
                Next m
            End If
            End If
            
           
           
            
            rs1.MoveNext
        Wend
     
        popuplistModel.Show
End If
rs1.close
End Sub

Sub addData(ST As String)
If RS.State = 1 Then RS.close


Select Case ST

Case "Author"

frmbook.txtWriter.Clear
RS.Open "Select * from MasterTbl where category='Author'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
  frmbook.txtWriter.AddItem RS!Name
   RS.MoveNext
Wend

Case "typesetter"

frmbook.txtTypeSetter.Clear
RS.Open "Select * from MasterTbl where category='typesetter'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
  frmbook.txtTypeSetter.AddItem RS!Name
   RS.MoveNext
Wend

Case "negative"

frmbook.txtNegativeby.Clear
RS.Open "Select * from MasterTbl where category='negative'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
  frmbook.txtNegativeby.AddItem RS!Name
   RS.MoveNext
Wend

Case "Lemination"

frmbook.cboLemination.Clear
RS.Open "Select * from MasterTbl where category='Lemination'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
  frmbook.cboLemination.AddItem RS!Name
   RS.MoveNext
Wend

Case "class"

frmbook.cboClass.Clear
RS.Open "Select * from MasterTbl where category='class'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
  frmbook.cboClass.AddItem RS!Name
   RS.MoveNext
Wend

Case "bkpart"

  frmbook.txtHead1.Clear
  frmbook.txtHead2.Clear
  frmbook.txtHead3.Clear
  frmbook.txtHead4.Clear
  frmbook.txtHead5.Clear
  frmbook.txtHead6.Clear
  frmbook.txtHead7.Clear
  frmbook.txtHead8.Clear
  
RS.Open "Select * from MasterTbl where category='bkpart'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
  frmbook.txtHead1.AddItem RS!Name
  frmbook.txtHead2.AddItem RS!Name
  frmbook.txtHead3.AddItem RS!Name
  frmbook.txtHead4.AddItem RS!Name
  frmbook.txtHead5.AddItem RS!Name
  frmbook.txtHead6.AddItem RS!Name
  frmbook.txtHead7.AddItem RS!Name
  frmbook.txtHead8.AddItem RS!Name
   RS.MoveNext
Wend

End Select


End Sub
Sub popuplist_client(ByVal ST As String, ByRef cn1 As ADODB.Connection, Optional ar As Integer, Optional font2 As String)

On Error Resume Next

Set rs1 = New ADODB.Recordset

rs1.Open ST, cn1, adOpenStatic, adLockReadOnly

'Exit Sub

If font2 = "h" Then
    Dim fill As New ADODB.Recordset
    Set fill = New ADODB.Recordset
    
    'popuplistVS.vs.FontSize = 12
    
    fill.Open ST, cn1
    
    Set popuplistVS.vs.DataSource = fill
    'popuplistvs.vs.Col = 2
    '-------------------------------
    
        For I = 0 To popuplistVS.vs.rows - 1
               DoEvents
            If font2 = "h" Then
               DoEvents
               DoEvents
               popuplistVS.vs.Cell(flexcpFontName, I, 1) = "Kundli"
               popuplistVS.vs.Cell(flexcpFontSize, I, 1) = 15
               
            End If
        Next
        
    '----------------------------------
    
     popuplistVS.vs.ColWidth(0) = 1500
     popuplistVS.vs.ColWidth(1) = 2800
     popuplistVS.vs.ColWidth(2) = 1600
     popuplistVS.vs.ColWidth(3) = 1600
     popuplistVS.vs.ColWidth(4) = 1600

    
    popuplistVS.Show
    Exit Sub
End If


If rs1.RecordCount > 0 Then
         If ar = 0 Then
            ar = rs1.Fields.Count
        ElseIf ar > rs1.Fields.Count Then
            ar = rs1.Fields.Count
        End If
        ReDim Array1(ar) As Variant
        For q = 0 To ar - 1
            Array1(q) = 0
        Next q
        Unload popuplistModel
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        
            For I = 1 To ar
            popuplistModel.ListView1.ColumnHeaders.Add I, , Trim(rs1.Fields(I - 1).Name)
            
            
            'If I = 1 Then
            '   popuplistModel.ListView1.ColumnHeaders(I).Width = 4500
            'Else
            '   popuplistModel.ListView1.ColumnHeaders(I).Width = 2000
            'End If
            
             
           Select Case searchType
            
           Case "party":
                
                bill = "0,5500,2500,2500,2500"
                a_strResult = Split(bill, ",")
                popuplistModel.ListView1.ColumnHeaders(I).Width = a_strResult(I)

           Case "ledger":
           
                bill = "0,5000,1500,3500,3000"
                a_strResult = Split(bill, ",")
                popuplistModel.ListView1.ColumnHeaders(I).Width = a_strResult(I)
           
           
           Case 3:
           
           Case 4:
            
           End Select
        
        Next I
        
        GridWidthSet = 0
        
        
        
         If font2 = "h" Then
        'If frmbook.bfont = "h" Then
            popuplistModel.ListView1.Font = hindi
            popuplistModel.ListView1.Font.Size = 15

        Else
            popuplistModel.ListView1.Font = english
            popuplistModel.ListView1.Font.Size = 12
        End If

        
        
        
        popuplistModel.ListView1.View = lvwReport
        Dim LItem As ListItem
        While Not rs1.EOF
        If (Not IsNull(rs1.Fields(0)) And rs1.Fields(0) <> "") Then
            
            Set LItem = popuplistModel.ListView1.ListItems.Add(, , rs1.Fields(0).value)
            
            If Len(rs1.Fields(0).value) > Len(rs1.Fields(0).Name) Then
               Array1(0) = Len(rs1.Fields(0).value)
            Else
               Array1(0) = Len(rs1.Fields(0).Name)
            End If
            If ar > 0 Then
                For m = 1 To ar - 1
                   If Not IsNull(rs1.Fields(m).value) Then LItem.SubItems(m) = rs1.Fields(m).value
                    If Len(rs1.Fields(m).value) > Len(rs1.Fields(m).Name) Then
                       Array1(m) = Len(rs1.Fields(m).value)
                    Else
                       Array1(m) = Len(rs1.Fields(m).Name)
                    End If
                    
                  
                Next m
            End If
            End If
            
            
            
            rs1.MoveNext
        Wend
     
        popuplistModel.Show
End If
rs1.close
End Sub
Public Function val_int(I As textbox, J As Integer) As Boolean

Dim a As Boolean
If J >= 48 And J <= 57 Or J = 8 Or J = 13 Or J = 46 Then
  val_int = True
Else
  val_int = False
End If


End Function



