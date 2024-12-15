VERSION 5.00
Begin VB.Form frmAuditTrailLog_Rem 
   Caption         =   "Enter Remark ...."
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   7680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   2400
   ScaleWidth      =   7680
   Begin VB.TextBox txtrem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   432
      TabIndex        =   0
      Top             =   684
      Width           =   6384
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   492
      Left            =   468
      TabIndex        =   1
      Top             =   1368
      Width           =   1572
   End
End
Attribute VB_Name = "frmAuditTrailLog_Rem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub addVoucherLog()

Dim rss As New ADODB.Recordset
Dim rss1 As New ADODB.Recordset
Dim vnumber As String

v_Remarks = UCase(txtrem.text)


If rss1.State = 1 Then rss1.close
rss1.Open "select top 1 [SubLedger],[Amount] from tmpVOUCHERS_del " & _
" where (vouchertype='" + Trim(vtypeNew) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate_ & "',103) And VoucherNumber = " & Trim(vno_) & ") order by vsno", con
If rss1.EOF = False Then
actionType_ = "Delete"
End If

        
If actionType_ <> "Delete" Then
'==========================================================================
'==========================================================================
        
        If rs1.State = 1 Then rs1.close
        rs1.Open "select * from  AuditTrail_Log  where (ActionType='Insert' and VoucherType='" + Trim(vtypeNew) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate_ & "',103)) And Vouchernumber = " + Trim(vno_), con
        If rs1.EOF = True Then
        
            If rss.State = 1 Then rss.close
            rss.Open "select * from  VOUCHERS_main  where vouchertype='" + Trim(vtypeNew) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate_ & "',103) And Vouchernumber = " + Trim(vno_), con
            If rss.EOF = False Then
                con.Execute "insert into AuditTrail_Log(VoucherID,VoucherType,ActionType,VoucherDate,VoucherNumber,[Description],Amount,ReasionForEdit,UserName) " & _
                 " values ('" & rss!VoucherID & "','" & rss!VoucherType & "','" & "Insert" & "','" & Format(rss!voucherDATE, "MM/dd/yyyy") & "','" & rss!VOUCHERNUMBER & "','" & rss!Particular & "','" & rss!amount & "','" & ReasionForEdit & "','" & UserName & "')"
            End If
        
        Else
        
            If rss.State = 1 Then rss.close
            rss.Open "select * from  VOUCHERS_main  where vouchertype='" + Trim(vtypeNew) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate_ & "',103) And Vouchernumber = " + Trim(vno_), con
            If rss.EOF = False Then
              con.Execute "update AuditTrail_Log set ReasionForEdit='" & v_Remarks & "', amount=" & rss!amount & ",amount_last=amount,dates='" & Format(Date, "MM/dd/yyyy") & "'  where (ActionType='Insert' and VoucherType='" + Trim(vtypeNew) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate_ & "',103)) And Vouchernumber = " + Trim(vno_)
            End If
        
        End If
        
        
        If v_vtype = "" Then Exit Sub
        
        If rss.State = 1 Then rss.close
        rss.Open "select * from VOUCHERS_bk where vouchertype='" + Trim(v_vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & v_vdate & "',103) And VoucherNumber = " + v_vnumber, con
        If rss.EOF = False Then
        
        v1 = Val(Trim(vno_))
        
            If Not (rss!VoucherType = v_vtype And rss!voucherDATE = vdate_ And rss!VOUCHERNUMBER = v1) Then
            
            
            If rss1.State = 1 Then rss1.close
            rss1.Open "select * from VOUCHERS_main where vouchertype='" + Trim(v_vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & v_vdate & "',103) And VoucherNumber = " + v_vnumber, con
            If rss1.EOF = False Then
               vnumber = rss1!VoucherID
            End If
            
            
            If rss1.State = 1 Then rss1.close
            rss1.Open "select * from VOUCHERS where vouchertype='" + Trim(vtypeNew) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate_ & "',103) And VoucherNumber = " + Trim(vno_), con
            While rss1.EOF = False
            
            If Len(rss1!subledger) > 0 Then
            h1 = rss1!subledger
            Else
            h1 = rss1!Genledger
            End If
            
            
            con.Execute "insert into AuditTrail_Log(VoucherID,VoucherType,ActionType,VoucherDate,VoucherNumber,[Description],Amount,ReasionForEdit,UserName,Narr,billno,billDate) " & _
             " values ('" & vnumber & "','" & rss1!VoucherType & "','" & "Edit" & "','" & Format(rss1!voucherDATE, "MM/dd/yyyy") & "','" & rss1!VOUCHERNUMBER & "','" & h1 & "','" & rss1!amount & "','" & v_Remarks & "','" & UserName & "','" & rss1!DESCRIPTION & "','" & rss1!CBND & "','" & rss1!billdate & "')"
            
            
            rss1.MoveNext
            Wend
        
        
        '==========================================
        Else
        '==========================================
        
            If rss1.State = 1 Then rss1.close
            rss1.Open "select * from VOUCHERS_main where vouchertype='" + Trim(v_vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & v_vdate & "',103) And VoucherNumber = " + v_vnumber, con
            If rss1.EOF = False Then
               vnumber = rss1!VoucherID
            End If
            
            Dim gledger, sledger, amt, billno, billdate, NARR, id As String
            gledger = ""
            sledger = ""
            amt = 0
            billno = 0
            billdate = ""
            NARR = ""
            id = ""
            
            
        
            For k1 = 1 To Voucherform.grid1.rows - 1 Step 2
            
            If Voucherform.grid1.TextMatrix(k1, 0) <> "" Then
            gledger = Voucherform.grid1.TextMatrix(k1, 0)
            sledger = Voucherform.grid1.TextMatrix(k1, 1)
            amt = IIf(Voucherform.grid1.TextMatrix(k1, 2) = "", Voucherform.grid1.TextMatrix(k1, 3), Voucherform.grid1.TextMatrix(k1, 2))
            billno = Voucherform.grid1.TextMatrix(k1, 4)
            
            billdate = Voucherform.grid1.TextMatrix(k1, 5)
            NARR = Voucherform.grid1.TextMatrix(k1 + 1, 0)
            id = Voucherform.grid1.TextMatrix(k1 + 1, 7)
            
            If (billdate = "__/__/____") Then
            billdate = ""
            End If
            
            
            If rss1.State = 1 Then rss1.close
            rss1.Open "select * from VOUCHERS_bk where vsno = '" & Trim(id) & "' order by Dates desc,vsno", con
            If rss1.EOF = False Then
                 
                dt1 = IIf(IsNull(rss1!billdate), "", rss1!billdate & "")
                
                
                If Not (gledger = rss1!Genledger And sledger = rss1!subledger And Val(amt) = rss1!amount And NARR = rss1!DESCRIPTION And billno = rss1!CBND And (billdate = dt1 Or billdate = dt1)) Then
                
                    If Len(rss1!subledger) > 0 Then
                      h1 = rss1!subledger
                    Else
                      h1 = rss1!Genledger
                    End If
                    
                    con.Execute "insert into AuditTrail_Log(VoucherID,VoucherType,ActionType,VoucherDate,VoucherNumber,[Description],Amount,ReasionForEdit,UserName,Narr,billno,billDate) " & _
                     " values ('" & vnumber & "','" & rss1!VoucherType & "','" & "Edit" & "','" & Format(rss1!voucherDATE, "MM/dd/yyyy") & "','" & rss1!VOUCHERNUMBER & "','" & h1 & "','" & amt & "','" & v_Remarks & "','" & UserName & "','" & NARR & "','" & billno & "','" & billdate & "')"
            
                
                End If
            
            Else
            
               If Len(Voucherform.grid1.TextMatrix(k1, 1)) > 0 Then
                 h1 = Voucherform.grid1.TextMatrix(k1, 1)
               Else
                 h1 = Voucherform.grid1.TextMatrix(k1, 0)
               End If
               
               
        
               con.Execute "insert into AuditTrail_Log(VoucherID,VoucherType,ActionType,VoucherDate,VoucherNumber,[Description],Amount,ReasionForEdit,UserName,Narr,billno,billDate) " & _
               " values ('" & vnumber & "','" & Voucherform.vtype & "','" & "Add" & "','" & Format(Voucherform.vdate, "MM/dd/yyyy") & "','" & Voucherform.vno & "','" & h1 & "','" & amt & "','" & v_Remarks & "','" & UserName & "','" & NARR & "','" & billno & "','" & billdate & "')"
            
            
            End If
            
            End If
            Next
            
        End If
        End If
    
''===========================================================================
Else
''===========================================================================
    
If rss1.State = 1 Then rss1.close
rss1.Open "select distinct [VoucherType],[VoucherDate] ,[VoucherNumber],[GenLedger],[SubLedger],[Amount],[CBND],[DESCRIPTION],[vsno],billdate from tmpVOUCHERS_del where (vouchertype='" + Trim(vtypeNew) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate_ & "',103) And VoucherNumber = " & Trim(vno_) & ") order by vsno", con
While rss1.EOF = False

If rs1.State = 1 Then rs1.close
rs1.Open "select * from VOUCHERS_main where vouchertype='" + Trim(v_vtype) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & v_vdate & "',103) And VoucherNumber = " + v_vnumber, con
If rs1.EOF = False Then
   vnumber = rs1!VoucherID
End If


    If Len(rss1!subledger) > 0 Then
    h1 = rss1!subledger
    Else
    h1 = rss1!Genledger
    End If

    con.Execute "insert into AuditTrail_Log(VoucherID,VoucherType,ActionType,VoucherDate,VoucherNumber,[Description],Amount,ReasionForEdit,UserName,Narr,billno,billDate) " & _
     " values ('" & vnumber & "','" & rss1!VoucherType & "','" & "Delete" & "','" & Format(rss1!voucherDATE, "MM/dd/yyyy") & "','" & rss1!VOUCHERNUMBER & "','" & h1 & "','" & rss1!amount & "','" & v_Remarks & "','" & UserName & "','" & rss1!DESCRIPTION & "','" & rss1!CBND & "','" & rss1!billdate & "')"


rss1.MoveNext
Wend

con.Execute "delete from tmpVOUCHERS_del where (vouchertype='" + Trim(vtypeNew) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate_ & "',103) And VoucherNumber = " & Trim(vno_) & ")"

'con.Execute "update VOUCHERS_bk set EntryNumber=null where EntryNumber=1 and vouchertype='" + Trim(vtypeNew) + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & vdate_ & "',103) And VoucherNumber = " & Trim(vno_) & ""


End If

'=========================================

'End If
'End If

v_Remarks = ""

End Sub
Private Sub cmdok_Click()
    
    
    
    If vtype1_ = "v" Then
       'AuditTrail_Log "V", actionType_, txtrem.text, vdate_, vno_, vtypeNew
       addVoucherLog   'New code
    ElseIf (vtype1_ = "I") Then
       AuditTrail_Log "I", actionType_, txtrem.text, vdate_, invoice.I_NO.text, vtypeNew
    ElseIf (vtype1_ = "CI") Then
       AuditTrail_Log "CI", actionType_, txtrem.text, vdate_, Critnote.I_NO.text, vtypeNew
    ElseIf (vtype1_ = "CM") Then
       AuditTrail_Log "CM", actionType_, txtrem.text, vdate_, countersale.I_NO.text, vtypeNew
   
    ElseIf (vtype1_ = "D") Then
       AuditTrail_Log "D", actionType_, txtrem.text, vdate_, Debitnotefile.TCNN.text, vtypeNew
    ElseIf (vtype1_ = "C") Then
       AuditTrail_Log "C", actionType_, txtrem.text, vdate_, Creditnotefile.TCNN.text, vtypeNew


    End If
    
    
    vtypes_ = ""
    vno_ = ""
    
    Unload Me
    
End Sub

Private Sub Form_Load()

Me.top = 500
Me.Left = 400

'txtrem.SetFocus

End Sub
Private Sub txtrem_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

cmdok_Click
Unload Me

End If

End Sub
