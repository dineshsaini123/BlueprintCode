VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form GLEDGERPRINT 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3450
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "GLPRINT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1_genshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show (New)"
      Height          =   675
      Left            =   1980
      Picture         =   "GLPRINT.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2205
      Width           =   1605
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   3075
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0078CFE9&
      Caption         =   "Crystal Report"
      Height          =   405
      Left            =   5265
      TabIndex        =   9
      Top             =   3375
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0078CFE9&
      Caption         =   "Dos Report"
      Height          =   390
      Left            =   5265
      TabIndex        =   8
      Top             =   3420
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   675
      Left            =   3660
      Picture         =   "GLPRINT.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1605
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   675
      Left            =   1980
      Picture         =   "GLPRINT.frx":17D4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1605
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   1995
      TabIndex        =   1
      Top             =   870
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   1965
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3540
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   4395
      TabIndex        =   2
      Top             =   870
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   315
      Left            =   3495
      TabIndex        =   7
      Top             =   930
      Width           =   465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   315
      Left            =   375
      TabIndex        =   6
      Top             =   870
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gen. Ledger Desc."
      Height          =   285
      Left            =   405
      TabIndex        =   5
      Top             =   360
      Width           =   1425
   End
End
Attribute VB_Name = "GLEDGERPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As Recordset
Sub GenOpening()
   
    Dim Trptrs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim rs3cash As New ADODB.Recordset
        
    Dim rs4 As New ADODB.Recordset
    Dim rs4c As New ADODB.Recordset
    Dim rs5 As New ADODB.Recordset
    
    Dim rs5C As New ADODB.Recordset
    Dim rs6 As New ADODB.Recordset
    Dim rs6C As New ADODB.Recordset
    Dim rs7 As New ADODB.Recordset
    Dim rs7C As New ADODB.Recordset
    Dim rs8 As New ADODB.Recordset
    Dim pur As New ADODB.Recordset
        
    
    Dim Ors1 As New ADODB.Recordset
    Dim Ors2 As New ADODB.Recordset
    Dim Ors3 As New ADODB.Recordset
    Dim Ors3cash As New ADODB.Recordset
    Dim Ors4 As New ADODB.Recordset
    Dim Ors41 As New ADODB.Recordset
    Dim Ors5 As New ADODB.Recordset
    Dim Ors51 As New ADODB.Recordset
    Dim Ors6 As New ADODB.Recordset
    Dim Ors61 As New ADODB.Recordset
    Dim Ors7 As New ADODB.Recordset
    Dim Ors71 As New ADODB.Recordset
    Dim Ors8 As New ADODB.Recordset
    
    Dim RSopBalDR As New ADODB.Recordset
    Dim RSopBalCR As New ADODB.Recordset
    Dim RsDRCR As New ADODB.Recordset
    
    Dim RsVDr As New ADODB.Recordset
    Dim RsVCr As New ADODB.Recordset
    Dim Balance As Double
    Dim OPBALANCE As Double
    Dim SDamount As Double
    Dim SCamount As Double
    Set rs1 = New ADODB.Recordset
    Dim viewsubledger As Boolean
    viewsubledger = False
    Balance = 0
    OPBALANCE = 0
    If RS.State = 1 Then
        RS.close
    End If
    
    
    con.Execute "Delete from TemprptTrialBalance"
    con.Execute "delete from treportGen"
 
 
    con.Execute "INSERT INTO TemprptTrialBalance (Gledger, OpeningBalance,userid,fyear,setupid)  SELECT Gledger.gledger, Gledger.YEAROPENING," & UId & " as Userid,'" & main.session & "'," & main.setupid & " from Gledger where  " & stringyear & " and Gledger='" & COMBOGENLEDGER.Text & "'"
    RsVDr.Open "SELECT GenLedger,sum(amount) as DAmount  FROM VOUCHERS   WHERE  " & stringyear & " and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='D' GROUP BY GenLedger ", con, adOpenDynamic, adLockReadOnly, adCmdText
    RsVCr.Open "SELECT GenLedger,sum(amount) as CAmount  FROM VOUCHERS  where   " & stringyear & " and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='C' GROUP BY GenLedger ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs1.Open "select GenLedger,  sum(Netamount) as SAmount from invoicea where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  GROUP BY GenLedger", con, adOpenDynamic, adLockReadOnly, adCmdText
    
    rs2.Open "select GenLedger,  sum(Netamount) as  SAmount from CREDITa where  " & stringyear & " and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  GROUP BY GenLedger ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs3.Open "select GenLedger,  sum(NETamount) as SAmount  from casha where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)and SUBLEDGER <>'CASH PARTY'  GROUP BY GenLedger ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs3cash.Open "select GenLedger,   sum(baa) as SAmount  from casha where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and SUBLEDGER <>'CASH PARTY'  GROUP BY GenLedger ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs4.Open "select pgld, sum(na) as SAmount  from Cnf1a where    " & stringyear & " and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  and dc ='D' GROUP BY pgld ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs4c.Open "select pgld, sum(na) as SAmount  from Cnf1a where   " & stringyear & " and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  and dc ='C' GROUP BY pgld ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs5.Open "select pgld, sum(na) as SAmount from dnfa where  " & stringyear & " and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  and dc ='D'  GROUP BY pgld ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs5C.Open "select pgld, sum(na) as SAmount from dnfa where   " & stringyear & " and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and  dc ='C'  GROUP BY pgld ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs6.Open "select gld, sum(a) as SAmount from Cnf1B where   " & stringyear & " and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and  dc ='D'  GROUP BY gld ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs6C.Open "select gld, sum(a) as SAmount from Cnf1B where   " & stringyear & " and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and dc ='C' GROUP BY gld ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs7.Open "select gld, sum(a) as SAmount from dnfB where    " & stringyear & " and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and dc ='D' GROUP BY gld ", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs7C.Open "select gld, sum(a) as SAmount from dnfB where    " & stringyear & " and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)and dc ='C' GROUP BY gld ", con, adOpenDynamic, adLockReadOnly, adCmdText
    'pur.Open "select GenLedger,  sum(Netamount) as SAmount from Purchasea where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  GROUP BY GenLedger", CON, adOpenDynamic, adLockReadOnly, adCmdText
    
    
    RS.Open "select * from TemprptTrialBalance where  " & stringyear & "", con, adOpenDynamic, adLockReadOnly, adCmdText
    
    a = con.Execute("Select sum(Gamount) as aa from InvoiceA  WHERE   " & stringyear & " and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)")(0).value
    If IsNull(a) Then a = 0
    con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & a & ")  where upper(Gledger) = 'SALES' and " & stringyear & " and userid=" & main.UId
    
    'a = CON.Execute("Select sum(Gamount) as aa from Purchasea  WHERE  convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)")(0).Value
    'If IsNull(a) Then a = 0
    'CON.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & a & ")  where Gledger = 'PURCHASE A/C'"
    
    
    a = con.Execute("Select sum(Gamount) as aa from CreditA where  " & stringyear & " and  convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)")(0).value
    If IsNull(a) Then a = 0
    
    con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & a & ")   where  " & stringyear & " and upper(Gledger) = 'SALES RETURN' and userid=" & main.UId
    a = con.Execute("Select sum(Gamount) as aa from CashA  where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)")(0).value
    
    If IsNull(a) Then a = 0
    
    con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & a & ")   where  " & stringyear & " and userid=" & main.UId & " and upper(Gledger) = 'SALES'"
    a = con.Execute("Select sum(BAA) as aa from CashA  where  convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)")(0).value
    If IsNull(a) Then a = 0
    con.Execute "Update TemprptTrialBalance set DAmount = (dAmount + " & a & ")   where  " & stringyear & " and userid=" & main.UId & " and upper(Gledger) = 'CASH-IN-HAND'"
    
    Dim tRS1 As New ADODB.Recordset
    
    '*******For Invoicec
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from INVOICEC  where  " & stringyear & " and Debitorcredit='Debit' AND convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND  upper(GENledger)='" & COMBOGENLEDGER.Text & "'  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    
    If tRS1.RecordCount > 0 Then
       While Not tRS1.EOF
                con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and upper(Gledger)='" & COMBOGENLEDGER.Text & "'"
          tRS1.MoveNext
       Wend
    End If
    If tRS1.State = 1 Then tRS1.close
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from INVOICEC where  " & stringyear & " and Debitorcredit='Credit'  AND convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND  upper(GENledger)='" & COMBOGENLEDGER.Text & "' GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
        While Not tRS1.EOF
                    con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")   where  " & stringyear & " and userid=" & main.UId & " and upper(Gledger)='" & COMBOGENLEDGER.Text & "'"
            tRS1.MoveNext
        Wend
    End If
    
    '******* End of INVOICEC
    
    
''    '*******For purchasec
''    If tRS1.State = 1 Then tRS1.close
''    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from Purchasec  where  " & stringyear & " and DEBITORCREDIT='Credit' AND convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND  upper(genledger)='" & COMBOGENLEDGER.Text & "'  GROUP BY GENLEDGER", CON, adOpenStatic, adLockReadOnly
''    If tRS1.RecordCount > 0 Then
''       While Not tRS1.EOF
''           CON.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")   where " & stringyear & " and userid=" & main.UId & " and Gledger='" & COMBOGENLEDGER.Text & "'"
''           tRS1.MoveNext
''       Wend
''    End If
    
'''    If tRS1.State = 1 Then tRS1.close
'''    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from Purchasec where " & stringyear & " and Debitorcredit='Debit'  AND convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)AND  upper(genledger)='" & COMBOGENLEDGER.Text & "' GROUP BY GENLEDGER", CON, adOpenStatic, adLockReadOnly
'''    If tRS1.RecordCount > 0 Then
'''        While Not tRS1.EOF
'''            CON.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")  where " & stringyear & " and userid=" & main.UId & " and Gledger='" & COMBOGENLEDGER.Text & "'"
'''            tRS1.MoveNext
'''        Wend
'''    End If
    
    '******* End of INVOICEC
    '*******For CrdeitC
    
    If tRS1.State = 1 Then tRS1.close
    
tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CreditC WHERE  " & stringyear & " and Debitorcredit='Debit' AND convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)AND  upper(genledger)='" & COMBOGENLEDGER.Text & "'  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
       While Not tRS1.EOF
 '         CON.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")  where Gledger='" & COMBOGENLEDGER.Text & "'"
          tRS1.MoveNext
       Wend
    End If
    If tRS1.State = 1 Then tRS1.close
        tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CreditC WHERE  " & stringyear & " and Debitorcredit='Credit' AND convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)AND  upper(genledger)='" & COMBOGENLEDGER.Text & "'   GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    
    If tRS1.RecordCount > 0 Then
        While Not tRS1.EOF
'            CON.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")  where Gledger='" & COMBOGENLEDGER.Text & "'"
            tRS1.MoveNext
        Wend
    End If
    
    '******* End of CreditC
    
    
    '*******For CashC
    If tRS1.State = 1 Then tRS1.close
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CashC WHERE  " & stringyear & " and Debitorcredit='Debit' AND convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  AND  upper(genledger)='" & COMBOGENLEDGER.Text & "'  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
       While Not tRS1.EOF
                con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")   where " & stringyear & " and userid=" & main.UId & " and Gledger='" & COMBOGENLEDGER.Text & "'"
          tRS1.MoveNext
       Wend
    End If
    If tRS1.State = 1 Then tRS1.close
tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CashC WHERE  " & stringyear & " and Debitorcredit='Credit' AND convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND  upper(genledger)='" & COMBOGENLEDGER.Text & "' GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
        While Not tRS1.EOF
            con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")   where " & stringyear & " and userid=" & main.UId & " and Gledger='" & COMBOGENLEDGER.Text & "'"
            tRS1.MoveNext
        
        
        Wend
    End If
    '*******For End of CashC
    If Not RS.BOF Then
           Do While Not RS.EOF
               OPBALANCE = RS!OpeningBalance
               SDamount = 0
               SCamount = 0
            If RsVDr.RecordCount > 0 Then
                  RsVDr.Find "genledger='" + Trim(RS!gledger) + "'"
                  If Not RsVDr.EOF Then
                        SDamount = SDamount + RsVDr!damount
                 End If
            End If

            If RsVCr.RecordCount > 0 Then
                  RsVCr.Find "genledger='" + Trim(RS!gledger) + "'"
                  If Not RsVCr.EOF Then
                        SCamount = SCamount + RsVCr!camount
                 End If
            End If

            If rs1.RecordCount > 0 Then
                  rs1.Find "genledger='" + Trim(RS!gledger) + "'"
                  If Not rs1.EOF Then
                      SDamount = SDamount + rs1!samount
                  End If
            End If
            
            If rs2.RecordCount > 0 Then
                  rs2.Find "genledger='" + Trim(RS!gledger) + "'"
                  If Not rs2.EOF Then
                      SCamount = SCamount + rs2!samount
                  End If
            End If
            
            If rs3.RecordCount > 0 Then
                  rs3.Find "genledger='" + Trim(RS!gledger) + "'"
                  If Not rs3.EOF Then
                      SDamount = SDamount + rs3!samount
                  End If
            End If
            
            If rs3cash.RecordCount > 0 Then
                  rs3cash.Find "genledger='" + Trim(RS!gledger) + "'"
                  If Not rs3cash.EOF Then
                      SCamount = SCamount + rs3cash!samount
                  End If
            End If
 
            
            If rs4.RecordCount > 0 Then
                  rs4.Find "pgld ='" + Trim(RS!gledger) + "'"
                  If Not rs4.EOF Then
                      SDamount = SDamount + rs4!samount
                  End If
            End If
            
            If rs4c.RecordCount > 0 Then
                  rs4c.Find "pgld= '" + RS!gledger + "'"
                  If Not rs4c.EOF Then
                      SCamount = SCamount + rs4c!samount
                  End If
            End If
            
            
            If rs5.RecordCount > 0 Then
                  rs5.Find "pgld='" + Trim(RS!gledger) + "'"
                  If Not rs5.EOF Then
                      SDamount = SDamount + rs5!samount
                  End If
            End If

            If rs5C.RecordCount > 0 Then
                  rs5C.Find "pgld='" + Trim(RS!gledger) + "'"
                  If Not rs5C.EOF Then
                      SCamount = SCamount + rs5C!samount
                  End If
            End If
            
            If rs6.RecordCount > 0 Then
                  rs6.Find "gld ='" + Trim(RS!gledger) + "'"
                  If Not rs6.EOF Then
                      SDamount = SDamount + rs6!samount
                  End If
            End If
            
            If rs6C.RecordCount > 0 Then
                  rs6C.Find "gld='" + Trim(RS!gledger) + "'"
                  If Not rs6C.EOF Then
                      SCamount = SCamount + rs6C!samount
                  End If
            End If
  
            If rs7.RecordCount > 0 Then
                  rs7.Find "gld='" + Trim(RS!gledger) + "'"
                  If Not rs7.EOF Then
                      SDamount = SDamount + rs7!samount
                  End If
            End If
            If rs7C.RecordCount > 0 Then
                  rs7C.Find "gld='" + Trim(RS!gledger) + "'"
                  If Not rs7C.EOF Then
                      SCamount = SCamount + rs7C!samount
                  End If
            End If
  
            con.Execute "UPDATE TemprptTrialBalance SET TemprptTrialBalance.DAmount = TemprptTrialBalance.DAmount+  " & SDamount & " ,TemprptTrialBalance.CAmount = TemprptTrialBalance.CAmount+  " & SCamount & "   where  " & stringyear & " and userid=" & main.UId & " and Gledger ='" & COMBOGENLEDGER.Text & "'"
            
            If RsVDr.RecordCount > 0 Then RsVDr.MoveFirst
            If RsVCr.RecordCount > 0 Then RsVCr.MoveFirst
            
            If rs1.RecordCount > 0 Then rs1.MoveFirst
            If rs2.RecordCount > 0 Then rs2.MoveFirst
            If rs3.RecordCount > 0 Then rs3.MoveFirst
            If rs3cash.RecordCount > 0 Then rs3cash.MoveFirst
            
            If rs4.RecordCount > 0 Then rs4.MoveFirst
            If rs4c.RecordCount > 0 Then rs4c.MoveFirst
            
            If rs5.RecordCount > 0 Then rs5.MoveFirst
            If rs5C.RecordCount > 0 Then rs5C.MoveFirst
            
            If rs6.RecordCount > 0 Then rs6.MoveFirst
            If rs6C.RecordCount > 0 Then rs6C.MoveFirst
            
            If rs7.RecordCount > 0 Then rs7.MoveFirst
            If rs7C.RecordCount > 0 Then rs7C.MoveFirst
            '''If pur.RecordCount > 0 Then pur.MoveFirst
            
            If Not RS.EOF Then
                RS.MoveNext
            End If
      
      Loop
    End If


End Sub

Private Sub COMBOGENLEDGER_LostFocus()
COMBOGENLEDGER = UCase(COMBOGENLEDGER)
    If Trim(COMBOGENLEDGER.Text) <> "" Then
                RS.Open "select * from gledger where " & stringyear & " and slf=0", con, adOpenStatic, adLockReadOnly, adCmdText
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
        RS.Open "select * from sledger where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.close
    Else
        Combosubledger.Text = ""
    End If
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command1_genshow_Click()


Screen.MousePointer = vbHourglass


If Trim(COMBOGENLEDGER.Text) <> "" Then
    
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date", , title1
        Exit Sub
    End If
    
    Commandshow.Enabled = False
    Command1_genshow.Enabled = False
    
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim rs5 As ADODB.Recordset
    Dim rs6 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Set rs3 = New ADODB.Recordset
    Set rs4 = New ADODB.Recordset
    Set rs5 = New ADODB.Recordset
    Set rs6 = New ADODB.Recordset
    Dim Balance As Double
    Balance = 0
    GenOpening
    Set rs1 = New ADODB.Recordset
    rs1.Open "select * from TemprptTrialBalance where  " & stringyear & "", con, adOpenDynamic, adLockReadOnly, adCmdText
    
    If rs1.RecordCount > 0 Then
         Balance = rs1!OpeningBalance + IIf(IsNull(rs1!damount), 0, rs1!damount) - IIf(IsNull(rs1!camount), 0, rs1!camount)
    End If
    Set rs1 = New ADODB.Recordset
    
    con.Execute "exec tmpGenledger  " & setupid & ",'" & session & "','" & Trim(COMBOGENLEDGER.Text) & "','" & date1.Text & "','" & date2.Text & "', " & UId & ""
    
    
   
    
    
    Dim header_, text_, Period_ As String
    headr_ = "GEN. LEDGER ACCOUNT"
    text_ = "** Opening Balance as on  " + Trim(date1.Text)
    Period_ = Trim(date1.Text) + "  To  " + Trim(date2.Text)
    
    If Balance > 0 Then
       con.Execute "insert into treportGen(header,Genledger,Text,aD,dorc,Period,userid,fyear,setupid,SUBLEDGER ) " & _
       " Values('" & headr_ & "','', '" & text_ & "'," & Balance & ",'D','" & Period_ & "'," & main.UId & ",'" & session & "','" & setupid & "','" & Trim(COMBOGENLEDGER.Text) & "')"
    Else
          con.Execute "insert into treportGen(header,Genledger,Text,aD,dorc,Period,userid,fyear,setupid,SUBLEDGER ) " & _
       " Values('" & headr_ & "','', '" & text_ & "'," & Balance & ",'C','" & Period_ & "'," & main.UId & ",'" & session & "','" & setupid & "','" & Trim(COMBOGENLEDGER.Text) & "')"
    End If
    
 
'making the balance in the output file
'If rs1.State = 1 Then rs1.Close
'rs1.Open "select * from treportGen  where " & stringyear & " order by vdate, vtype,vno", con, adOpenStatic, adLockOptimistic, adCmdText
Set rs1 = con.Execute("exec fatch_treportGen")

    Balance = 0
    If Not rs1.BOF Then
        pb.value = 0
        pb.Max = rs1.RecordCount
        rs1.MoveFirst
        Do While Not rs1.EOF
            Balance = Balance + IIf(IsNull(rs1!ad), 0, rs1!ad) - IIf(IsNull(rs1!aC), 0, rs1!aC)
            'rs1!Balance = Round(Balance, 2)
            'rs1.update
            'aaa = rs1!sno
            'aaa = RS!sno
            
            con.Execute "update treportGen set Balance = " & Round(Balance, 2) & " where userid=" & main.UId & " and sno=" & rs1!sno
            
            If Not rs1.EOF Then
                rs1.MoveNext
                pb.value = pb.value + 1
            End If
        Loop
    End If

If RS.State = 1 Then RS.close
 
 main.reportname = "Gen. Ledger"
 viewgenledger.genreport
 PrintOption.Show

Else
    MsgBox "gen. ledger not selected", , title1
End If



Commandshow.Enabled = True
Command1_genshow.Enabled = True
Screen.MousePointer = vbDefault


End Sub

Private Sub CommandReturn_Click()
    Unload Me
End Sub
Private Sub Commandshow_Click()
If Trim(COMBOGENLEDGER.Text) <> "" Then
    
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date", , title1
        Exit Sub
    End If
    
    Commandshow.Enabled = False
    
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim rs5 As ADODB.Recordset
    Dim rs6 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Set rs3 = New ADODB.Recordset
    Set rs4 = New ADODB.Recordset
    Set rs5 = New ADODB.Recordset
    Set rs6 = New ADODB.Recordset
    Dim Balance As Double
    Balance = 0
    GenOpening
    Set rs1 = New ADODB.Recordset
    rs1.Open "select * from TemprptTrialBalance where  " & stringyear & "", con, adOpenDynamic, adLockReadOnly, adCmdText
    
    If rs1.RecordCount > 0 Then
         Balance = rs1!OpeningBalance + IIf(IsNull(rs1!damount), 0, rs1!damount) - IIf(IsNull(rs1!camount), 0, rs1!camount)
    End If
    Set rs1 = New ADODB.Recordset
    
    
     rs1.Open "Select * from treportGen where " & stringyear, con, adOpenDynamic, adLockPessimistic
    
    rs1.AddNew
    rs1!header = "GEN. LEDGER ACCOUNT"
    rs1!Genledger = " "
    rs1!Text = "** Opening Balance as on  " + Trim(date1.Text)
    If Balance > 0 Then
       rs1!ad = Balance
       rs1!dorc = "D"
    ElseIf Balance < 0 Then
       rs1!aC = Abs(Balance)
       rs1!dorc = "C"
    End If
    rs1!Period = Trim(date1.Text) + "  To  " + Trim(date2.Text)
    rs1!Genledger = Trim(COMBOGENLEDGER.Text)
        rs1!userid = main.UId
    rs1!fyear = main.session
    rs1!setupid = main.setupid
    rs1!SUBLEDGER = Trim(COMBOGENLEDGER.Text)
    rs1.UpdateBatch
   
   
   
   '************* voucher start
   If RS.State = 1 Then RS.close
    RS.Open "select * from vouchers where  " & stringyear & " and upper(genledger)='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by VoucherDate", con, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not RS.BOF Then
        Do While Not RS.EOF
             rs1.AddNew
             rs1!Genledger = Trim(RS!Genledger)
             rs1!SUBLEDGER = Trim(RS!SUBLEDGER)
             rs1!vdate = RS!voucherDATE & ""
             rs1!vtype = Trim(RS!vouchertype)
             rs1!vno = Trim(Str(RS!VOUCHERNUMBER))
             If Len(Trim(RS!DESCRIPTION)) = 0 Then
                rs1!narration = " "
             Else
                rs1!narration = Trim(RS!DESCRIPTION)
             End If
             If Trim(UCase(RS!DebitorCredit)) = Trim(UCase("D")) Then
                rs1!ad = RS!amount
                rs1!dorc = "D"
            '   balance = balance + rs!amount
            '   rs1!balance = balance
             Else
                rs1!aC = RS!amount
                rs1!dorc = "C"
             '  balance = balance - rs!amount
             '  rs1!balance = balance
            End If
            If Len(Trim(RS!cbnd)) = 0 Or IsNull(RS!cbnd) Then
               rs1!cbno = " "
            Else
               rs1!cbno = RS!cbnd & ""
            End If
               rs1!userid = main.UId
               rs1!fyear = main.session
               rs1!setupid = main.setupid
               rs1.update
            If Not RS.EOF Then
               RS.MoveNext
            End If
        Loop
    End If
'***************   voucher end
'selecting the invoice end part data
'sales GL start
 If UCase(Trim(COMBOGENLEDGER.Text)) = "SALES" Then
    rs3.Open "select * from invoiceA where  " & stringyear & " and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by invoiceno", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs3.EOF Then
    pb.Max = rs3.RecordCount
    
           Do While Not rs3.EOF
              If rs3!gamount <> 0 Then
                 rs1.AddNew
                 rs1!Genledger = "SALES"
                 rs1!vdate = rs3!invoiceDate & ""
                 rs1!vtype = Trim("I")
                 rs1!vno = Trim(rs3!invoiceNo)
                 rs1!narration = "Sales"                '"Sales Invoice"
                 balance1 = rs3!gamount
                 If rs2.State = 1 Then rs2.close
                 rs2.Open "select * from invoiceC where  " & stringyear & " and upper(genledger)='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and invoiceno =" + Str(rs3!invoiceNo), con, adOpenStatic, adLockReadOnly, adCmdText
                 If rs2.RecordCount > 0 Then
                    Do While Not rs2.EOF
                       If Left(rs2!DebitorCredit, 1) = "D" Then
                          balance1 = balance1 - rs2!amount
                       Else
                          balance1 = balance1 + rs2!amount
                       End If
                       rs2.MoveNext
                    Loop
                 End If
                 rs1!aC = balance1
                 rs1!narration = "Sales"           '"Sales Invoice"
                 rs1!dorc = "C"
                 rs1!userid = main.UId
                 rs1!fyear = main.session
                 rs1!setupid = main.setupid
                 rs1.update
                 
             End If
             
             If pb.Max = pb.value Then
             pb.value = 0
             Else
             pb.value = pb.value + 1
             End If
             
             If Not rs3.EOF Then rs3.MoveNext
               
    
         Loop
    End If
    rs3.close
   '================================================
   
   
   
   
   
   
   '=================================================
   'CASH COUNTER SALE
    rs3.Open "select * from CASHA where  " & stringyear & " and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by invoiceno", con, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs3.EOF Then
       Do While Not rs3.EOF
          If rs3!gamount <> 0 Then
             rs1.AddNew
             rs1!Genledger = "SALES"
             rs1!vdate = rs3!invoiceDate & ""
             rs1!vtype = Trim("S")
             rs1!vno = Trim(rs3!invoiceNo)
             rs1!narration = "Sales C/M"
             balance1 = rs3!gamount
             If rs2.State = 1 Then rs2.close
             rs2.Open "select * from CASHC where  " & stringyear & " and upper(genledger)='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and invoiceno =" + Str(rs3!invoiceNo), con, adOpenStatic, adLockReadOnly, adCmdText
             
             If rs2.RecordCount > 0 Then
                Do While Not rs2.EOF
                   If Left(rs2!DebitorCredit, 1) = "D" Then
                      balance1 = balance1 - rs2!amount
                   Else
                      balance1 = balance1 + rs2!amount
                   End If
                   rs2.MoveNext
                Loop
             End If
             rs1!aC = balance1
             rs1!narration = "Sales C/M"
             rs1!dorc = "C"
             
             rs1!userid = main.UId
             rs1!fyear = main.session
             rs1!setupid = main.setupid
             rs1.update
             
             
          End If
          If Not rs3.EOF Then rs3.MoveNext
       Loop
    End If
    rs3.close
Else
    rs2.Open "select * from invoiceC where  " & stringyear & " and upper(genledger)='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    
    Do While Not rs2.EOF
       If rs2!amount <> 0 Then
          rs1.AddNew
          rs1!Genledger = rs2!Genledger & ""
          rs1!vdate = rs2!invoiceDate & ""
          rs1!vtype = Trim("I")
          rs1!narration = "Sales"                  '"Sales Invoice"
          rs1!vno = Trim(rs2!invoiceNo)
          If Left(rs2!DebitorCredit, 1) = "D" Then
             rs1!ad = rs2!amount
             rs1!dorc = "D"
          Else
            rs1!aC = rs2!amount
            rs1!dorc = "C"
          End If
          
          rs1!userid = main.UId
          rs1!fyear = main.session
          rs1!setupid = main.setupid
          rs1.update

          
      End If
      If Not rs2.EOF Then rs2.MoveNext
    Loop
    
    If rs2.State = 1 Then rs2.close
    'CASH COUNTER SALE
    rs2.Open "select * from CASHC where  " & stringyear & " and upper(genledger)='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    
    Do While Not rs2.EOF
       If rs2!amount <> 0 Then
          rs1.AddNew
          rs1!Genledger = rs2!Genledger & ""
          rs1!vdate = rs2!invoiceDate & ""
          rs1!vtype = Trim("S")
          rs1!narration = "Sales C/M"
          rs1!vno = Trim(rs2!invoiceNo)
          If Left(rs2!DebitorCredit, 1) = "D" Then
             rs1!ad = rs2!amount
             rs1!dorc = "D"
          Else
             rs1!aC = rs2!amount
             rs1!dorc = "C"
          End If
          
        rs1!userid = main.UId
        rs1!fyear = main.session
        rs1!setupid = main.setupid
        rs1.update
          
       End If
       If Not rs2.EOF Then rs2.MoveNext
    Loop
End If
'Sales RETURNS GL start
If UCase(Trim(COMBOGENLEDGER.Text)) = "SALES RETURN" Then
   rs4.Open "select * from CREDITA where  " & stringyear & " and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by invoiceno", con, adOpenStatic, adLockReadOnly, adCmdText
   
   If Not rs4.BOF Then
      Do While Not rs4.EOF
         If rs4!gamount <> 0 Then
            rs1.AddNew
            rs1!Genledger = "SALES RETURN"
            rs1!vdate = rs4!invoiceDate & ""
            rs1!vtype = Trim("C")
            rs1!vno = Trim(rs4!invoiceNo)
            rs1!narration = "Credit Note(I)"
            balance1 = rs4!gamount
            If rs2.State = 1 Then rs2.close
                   rs2.Open "select * from CREDITC where  " & stringyear & " and upper(genledger)='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and invoiceno =" + Str(rs4!invoiceNo), con, adOpenStatic, adLockReadOnly, adCmdText
                    
                    If rs2.RecordCount > 0 Then
                       Do While Not rs2.EOF
                            If Left(rs2!DebitorCredit, 1) = "D" Then
                                balance1 = balance1 + rs2!amount
                            Else
                                balance1 = balance1 - rs2!amount
                            End If
                            rs2.MoveNext
                        Loop
                    End If
                    rs1!ad = balance1
                    rs1!dorc = "D"
                  
                
                rs1!userid = main.UId
                rs1!fyear = main.session
                rs1!setupid = main.setupid
                rs1.update

                End If
                If Not rs4.EOF Then
                   rs4.MoveNext
                End If
            Loop
        End If
        rs4.close
    Else
       If rs2.State = 1 Then rs2.close
       rs2.Open "select * from CREDITC where  " & stringyear & " and upper(genledger)='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rs2.EOF
            If rs2!amount <> 0 Then
                rs1.AddNew
                rs1!Genledger = rs2!Genledger & ""
                rs1!narration = "Credit Note(I)"
                rs1!vdate = rs2!invoiceDate & ""
                rs1!vtype = Trim("I")
                rs1!vno = Trim(rs2!invoiceNo)
                If Left(rs2!DebitorCredit, 1) = "D" Then
                    rs1!ad = rs2!amount
                    rs1!dorc = "D"
'                   BALANCE = BALANCE + rs2!amount
                Else
                    rs1!aC = rs2!amount
                    rs1!dorc = "C"
 '                  BALANCE = BALANCE - rs2!amount
                End If
  '              rs1!BALANCE = BALANCE
  
            rs1!userid = main.UId
            rs1!fyear = main.session
            rs1!setupid = main.setupid
            rs1.update
  
            End If
            
            If Not rs2.EOF Then
               rs2.MoveNext
            End If
        Loop
    End If

'CREDIT NOTE
    rs3.Open "select * from Cnf1a where  " & stringyear & " and pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs3.BOF Then
        Do While Not rs3.EOF
            rs1.AddNew
            rs1!Genledger = rs3!Pgld & ""
            rs1!vdate = rs3!Cnd & ""
            rs1!vtype = "C"
            rs1!vno = Trim(Str(rs3!cnn))
            If Trim(UCase(rs3!dc)) = Trim(UCase("D")) Then
                rs1!ad = rs3!na
                rs1!dorc = "D"
            Else
                rs1!aC = rs3!na
                rs1!dorc = "C"
            End If
            rs1!narration = "Credit Note"
            rs1!userid = main.UId
            rs1!fyear = main.session
            rs1!setupid = main.setupid
            rs1.update
            If Not rs4.EOF Then
                rs3.MoveNext
            End If
        Loop
    End If
    'CREDIT NOTE B
   rs4.Open "select * from Cnf1b where  " & stringyear & " and gld ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs4.BOF Then
        Do While Not rs4.EOF
            rs1.AddNew
            rs1!Genledger = Trim(rs4!gld)
            rs1!vdate = rs4!Cnd & ""
            rs1!vtype = "C"
            rs1!vno = Trim(Str(rs4!cnn))
            If Trim(UCase(rs4!dc)) = Trim(UCase("D")) Then
                rs1!ad = rs4!a
                rs1!dorc = "D"
            Else
                rs1!aC = rs4!a
                rs1!dorc = "C"
            End If
            rs1!narration = "Credit Note"
               rs1!userid = main.UId
            rs1!fyear = main.session
            rs1!setupid = main.setupid
            rs1.update
            If Not rs4.EOF Then
                rs4.MoveNext
            End If
        Loop
    End If
'credit note b end
    
    'debit NOTE
     rs5.Open "select * from dnfa where  " & stringyear & " and pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs5.BOF Then
        Do While Not rs5.EOF
            rs1.AddNew
            rs1!Genledger = rs5!Pgld & ""
            rs1!vdate = rs5!dnd & ""
            rs1!vtype = "C"
            rs1!vno = Trim(Str(rs5!dnn))
            If Trim(UCase(rs5!dc)) = Trim(UCase("D")) Then
                rs1!ad = rs5!na
                rs1!dorc = "D"
            Else
                rs1!aC = rs5!na
                rs1!dorc = "C"
            End If
            rs1!narration = "Debit Note"
               rs1!userid = main.UId
            rs1!fyear = main.session
            rs1!setupid = main.setupid
          rs1.update
            If Not rs5.EOF Then
                rs5.MoveNext
            End If
        Loop
    End If
    'Debit NOTE B
   rs6.Open "select * from dnfb where " & stringyear & " and gld ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs6.BOF Then
        Do While Not rs6.EOF
            rs1.AddNew
            rs1!Genledger = rs6!gld & ""
            rs1!vdate = rs6!dnd & ""
            rs1!vtype = "C"
            rs1!vno = Trim(Str(rs6!dnn))
            If Trim(UCase(rs6!dc)) = Trim(UCase("D")) Then
                rs1!ad = rs6!a
                rs1!dorc = "D"
            Else
                rs1!aC = rs6!a
                rs1!dorc = "C"
            End If
            rs1!narration = "Debit Note"
            rs1!userid = main.UId
               rs1!fyear = main.session
            rs1!setupid = main.setupid
            rs1.update
            If Not rs6.EOF Then
                rs6.MoveNext
            End If
        Loop
    End If
'debit note b end
    
    
rs1.close

'making the balance in the output file
'rs1.Open "select * from treportGen where " & stringyear & " order by vdate, vtype,vno", CON, adOpenStatic, adLockOptimistic, adCmdText
rs1.Open "select * from treportGen  where " & stringyear & " order by vdate, vtype,vno", con, adOpenStatic, adLockOptimistic, adCmdText

    Balance = 0
    If Not rs1.BOF Then
        pb.value = 0
        pb.Max = rs1.RecordCount
        rs1.MoveFirst
        Do While Not rs1.EOF
            Balance = Balance + IIf(IsNull(rs1!ad), 0, rs1!ad) - IIf(IsNull(rs1!aC), 0, rs1!aC)
            rs1!Balance = Round(Balance, 2)
             rs1.update
            If Not rs1.EOF Then
                rs1.MoveNext
                pb.value = pb.value + 1
            End If
        Loop
    End If
    If rs1.State = 1 Then
        rs1.close
    End If
    If RS.State = 1 Then
        RS.close
    End If
    main.reportname = "Gen. Ledger"
    viewgenledger.genreport
    
'Unload PrintOption
'Load PrintOption

PrintOption.Show

Else
    MsgBox "gen. ledger not selected", , title1
End If

Commandshow.Enabled = True

End Sub

Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'SendKeys "{TAB}"
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
    'SendKeys "{TAB}"
End If
End Sub

Private Sub date2_LostFocus()
    If Trim(date2.Text) <> "" Then
        If Not checkdate(Trim(date2.Text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub Form_Load()

'Option1.Value = True

''''''''
''''''''Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
''''''''    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
''''''''        Unload VB.Screen.ActiveForm
''''''''    End If
''''''''Loop


Command1_genshow.Visible = False

If UserName = "v" Then
   Command1_genshow.Visible = True
End If


BackColorFrom Me

Me.Top = 0
Me.Left = 0
Set RS = New ADODB.Recordset
  RS.Open "select * from gledger where " & stringyear & " and slf=0", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            COMBOGENLEDGER.AddItem Trim(RS!gledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    RS.Open "Select * from setup1 where " & stringyear & "", con, adOpenStatic, adLockReadOnly
    CNSetup
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
    RS.close
End Sub

