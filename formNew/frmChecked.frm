VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChecked 
   Caption         =   "Voucher List For Check ..."
   ClientHeight    =   10392
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   16404
   Icon            =   "frmChecked.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10392
   ScaleWidth      =   16404
   Begin VB.CheckBox option_selectAll 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   10980
      TabIndex        =   8
      Top             =   216
      Width           =   1740
   End
   Begin VB.CommandButton CommandReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   14472
      Picture         =   "frmChecked.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   180
      Width           =   1152
   End
   Begin VB.CommandButton Commandsave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sa&ve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   13104
      Picture         =   "frmChecked.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   180
      Width           =   1284
   End
   Begin VB.ComboBox cbovtype 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6840
      TabIndex        =   4
      Top             =   180
      Width           =   3840
   End
   Begin MSComCtl2.DTPicker fDate_ 
      Height          =   336
      Left            =   180
      TabIndex        =   0
      Top             =   216
      Width           =   1668
      _ExtentX        =   2942
      _ExtentY        =   593
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
      Format          =   129892353
      CurrentDate     =   39795
   End
   Begin MSComCtl2.DTPicker tdate_ 
      Height          =   336
      Left            =   2664
      TabIndex        =   2
      Top             =   216
      Width           =   1668
      _ExtentX        =   2942
      _ExtentY        =   593
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
      Format          =   129892353
      CurrentDate     =   39795
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   9372
      Left            =   144
      TabIndex        =   7
      Top             =   936
      Width           =   15492
      _cx             =   27326
      _cy             =   16531
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777209
      ForeColor       =   16711680
      BackColorFixed  =   16777173
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777166
      BackColorAlternate=   16777209
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   420
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmChecked.frx":17D4
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2052
      TabIndex        =   3
      Top             =   252
      Width           =   624
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Type :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   4896
      TabIndex        =   1
      Top             =   216
      Width           =   2352
   End
End
Attribute VB_Name = "frmChecked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub search()

Screen.MousePointer = vbHourglass
   
   
   
   Dim value_ As String
   value_ = "(invoicedate>=convert(smalldatetime,'" & fDate_.value & "',103) and invoicedate<=convert(smalldatetime,'" & tdate_.value & "',103))"


   
   Dim vtype As String
   
   
 
 
    If (cbovtype.text = "Sale Invoice") Then
        str_ = "SELECT SUBLEDGER as PARTICULARS,INVOICEDATE as VDATE, 'Sales Invoice'  AS BType,INVOICENO as VNO," & _
        "NETAMOUNT as Amount,Checked_YesNo,CheckedBy from INVOICEA  where Checked_YesNo=0 and BAuthorized=1  and " & value_ & "  order  BY INVOICENO"
    
    ElseIf (cbovtype.text = "Counter Sale") Then
            
            str_ = "SELECT SUBLEDGER as PARTICULARS,INVOICEDATE as VDATE, 'Counter Sale'  AS BType,INVOICENO as VNO," & _
        "NETAMOUNT as Amount,Checked_YesNo,CheckedBy from cashA  where Checked_YesNo=0 and " & value_ & "  order  BY INVOICENO"
 
    ElseIf (cbovtype.text = "Credit Note") Then
        
        value_ = "(cnd>=convert(smalldatetime,'" & fDate_.value & "',103) and cnd<=convert(smalldatetime,'" & tdate_.value & "',103))"

        str_ = "SELECT PSLD as PARTICULARS,CND as VDATE, 'Credit Note'  AS BType,CNN as VNO," & _
        "NA as Amount,Checked_YesNo,CheckedBy from cnf1a  where Checked_YesNo=0 and BAuthorized=1  and " & value_ & "  order  BY cnn"
    
    ElseIf (cbovtype.text = "Credit Note (Item)") Then
       str_ = "SELECT SUBLEDGER as PARTICULARS,INVOICEDATE as VDATE, 'Credit Note (Item)'  AS BType,INVOICENO as VNO," & _
        "NETAMOUNT as Amount,Checked_YesNo,CheckedBy from CREDITA  where Checked_YesNo=0 and BAuthorized=1  and " & value_ & "  order  BY INVOICENO"
    ElseIf (cbovtype.text = "Debit Note") Then
       
       str_ = "SELECT SUBLEDGER as PARTICULARS,INVOICEDATE as VDATE, 'Debit Note'  AS BType,DNN as VNO," & _
        "Amount,Checked_YesNo,CheckedBy from debitRegister  where Checked_YesNo=0 and BAuthorized=1  and " & value_ & "  order  BY dnn"
    
    ElseIf (cbovtype.text = "Payment Voucher" Or cbovtype.text = "Receipt Voucher" Or cbovtype.text = "Journal Voucher") Then
        
        vtype_ = Left(cbovtype.text, 1)
        
        value_ = " (VoucherDate>=convert(smalldatetime,'" & fDate_.value & "',103) and VoucherDate<=convert(smalldatetime,'" & tdate_.value & "',103)) and Checked_YesNo=0 and vouchertype='" & vtype_ & "'"
        
        str_ = "SELECT Particular as PARTICULARS,VoucherDate as VDATE, '" & cbovtype.text & "' ,VoucherID as VNO," & _
              "Amount,Checked_YesNo,CheckedBy from VOUCHERS_Main   where  " + value_ + "" & _
              "order  BY VoucherID"
    End If
        

 
 
   Dim r1 As Integer
   vs.Clear

 
   
   
   r1 = 1
   
   vs.rows = 1
   vs.Cols = 7
   action_v = ""
 
   If rs1.State = 1 Then rs1.Close
   
  
   rs1.Open str_, con, adOpenDynamic, adLockOptimistic
   
   While rs1.EOF = False
   
   vs.rows = vs.rows + 1
   
   If rs1!Checked_YesNo = 0 Then
    action_v = 0
   Else
    action_v = 1
   End If
   
   vs.TextMatrix(r1, 0) = rs1!vdate
   vs.TextMatrix(r1, 1) = rs1!vno
   
   vs.TextMatrix(r1, 2) = rs1!Particulars
   vs.TextMatrix(r1, 3) = rs1!amount
   vs.TextMatrix(r1, 4) = rs1!CheckedBy & ""
   vs.TextMatrix(r1, 5) = action_v
   
   vs.TextMatrix(r1, 6) = rs1(2)
      
   
   
   r1 = r1 + 1
   rs1.MoveNext
   Wend
   
   
''==========================================================================================================
   
   
   
vs.WordWrap = True
   
vs.FormatString = "VOUCHER DATE|VNO|PARTICULARS|Amount|CREATED BY|AWAITING ACTION|VTYPE"
vs.ColWidth(0) = 1500
vs.ColWidth(1) = 1000
vs.ColWidth(2) = 5500
vs.ColWidth(3) = 1800
vs.ColWidth(4) = 1800
vs.ColWidth(5) = 1200
vs.ColWidth(6) = 1800


Screen.MousePointer = vbDefault



End Sub



Private Sub cbovtype_Click()
search
End Sub

Private Sub CommandReturn_Click()
Unload Me
End Sub

Private Sub Commandsave_Click()

saveData

End Sub
Sub saveData()
   
On Error GoTo ss:

Dim sysName_date As String

sysName_date = com_name & " - " & Date
   
   If MsgBox("Want to Set ?", vbQuestion + vbYesNo) = vbYes Then
        
   Screen.MousePointer = vbHourglass
   
   
   Dim din As Integer
         
   If cbovtype.text = "Sale Invoice" Then
        
        For J = 1 To vs.rows - 1
          
            If vs.TextMatrix(J, 5) <> "" Then
              
              If vs.TextMatrix(J, 5) = True Then
                 If vs.TextMatrix(J, 5) = True Then din = 1
                 con.Execute "update INVOICEA set Checked_YesNo=" & din & ",CheckedBy='" & UserName & "' where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
              
                 If (AuditTrail = "y") Then
                    
                    actionType_ = "Insert"
                    vtype1_ = "I"
                    vtypeNew = Left(vs.TextMatrix(J, 6), 1)
                    vdate_ = vs.TextMatrix(J, 0)
                    vno_ = vs.TextMatrix(J, 1)
    
                    AuditTrail_Log "I", "Insert", "", vdate_, vno_, vtypeNew
        
                 End If
              
              End If
            
            End If
          
        Next
        
    ElseIf (cbovtype.text = "Journal Voucher" Or cbovtype.text = "Receipt Voucher" Or cbovtype.text = "Payment Voucher") Then
        
        For J = 1 To vs.rows - 1
          
            If vs.TextMatrix(J, 5) <> "" Then
              
              If vs.TextMatrix(J, 5) = True Then
                 If vs.TextMatrix(J, 5) = True Then din = 1
                 con.Execute "update VOUCHERS_Main set Checked_YesNo=" & din & ",CheckedBy='" & UserName & "' where " & stringyear & " and VoucherID=" & vs.TextMatrix(J, 1) & ""
              
                 If (AuditTrail = "y") Then
                    
                    actionType_ = "Insert"
                    vtype1_ = "V"
                    vtypeNew = Left(vs.TextMatrix(J, 6), 1)
                    vdate_ = vs.TextMatrix(J, 0)
                    vno_ = vs.TextMatrix(J, 1)
    
                    AuditTrail_Log "V", "Insert", "", vdate_, vno_, vtypeNew
        
                 End If
    
              End If
            
            End If
          
        Next
        
  ElseIf cbovtype.text = "Credit Note (Item)" Then
  
        For J = 1 To vs.rows - 1
        
          If vs.TextMatrix(J, 5) = "True" Then vs.TextMatrix(J, 5) = -1
          If vs.TextMatrix(J, 5) = True Then
            If vs.TextMatrix(J, 5) = True Then din = 1
              con.Execute "update CREDITA set Checked_YesNo=" & din & ",CheckedBy='" & UserName & "' where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            
              If (AuditTrail = "y") Then
                    
                    actionType_ = "Insert"
                    vtype1_ = "CI"
                    vtypeNew = Left(vs.TextMatrix(J, 6), 1)
                    vdate_ = vs.TextMatrix(J, 0)
                    vno_ = vs.TextMatrix(J, 1)
    
                    AuditTrail_Log "CI", "Insert", "", vdate_, vno_, vtypeNew
        
             End If
            
          End If
          
        Next
  
  
  ElseIf cbovtype.text = "Counter Sale" Then
        
        For J = 1 To vs.rows - 1
        
        
         If vs.TextMatrix(J, 5) <> "" Then
          
          If vs.TextMatrix(J, 5) = True Then
             If vs.TextMatrix(J, 5) = True Then din = 1
             con.Execute "update casha set Checked_YesNo=" & din & ",CheckedBy='" & UserName & "' where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          
          
             If (AuditTrail = "y") Then
                    
                    actionType_ = "Insert"
                    vtype1_ = "CM"
                    vtypeNew = Left(vs.TextMatrix(J, 6), 1)
                    vdate_ = vs.TextMatrix(J, 0)
                    vno_ = vs.TextMatrix(J, 1)
    
                    AuditTrail_Log "CM", "Insert", "", vdate_, vno_, vtypeNew
        
             End If
          
          
          End If
        
        End If
      
       Next
        
  
  ElseIf cbovtype.text = "Credit Note" Then
  
     For J = 1 To vs.rows - 1
          
          If vs.TextMatrix(J, 5) = True Then
            
            If vs.TextMatrix(J, 5) = False Then
               vs.TextMatrix(J, 5) = 0
            Else
               vs.TextMatrix(J, 5) = -1
            End If
            
           
            con.Execute "update cnf1a set Checked_YesNo=" & vs.TextMatrix(J, 5) & ",CheckedBy='" & UserName & "'  where " & stringyear & " and cnn=" & vs.TextMatrix(J, 1) & ""
           
            
            If (AuditTrail = "y") Then
                    
                actionType_ = "Insert"
                vtype1_ = "C"
                vtypeNew = Left(vs.TextMatrix(J, 6), 1)
                vdate_ = vs.TextMatrix(J, 0)
                vno_ = vs.TextMatrix(J, 1)

                AuditTrail_Log "C", "Insert", "", vdate_, vno_, vtypeNew
        
            End If
           
          End If
        
     Next
  
  
  ElseIf cbovtype.text = "Debit Note" Then
  
        For J = 1 To vs.rows - 1
        
          If vs.TextMatrix(J, 5) = True Then
            
            If vs.TextMatrix(J, 5) = False Then
               vs.TextMatrix(J, 5) = 0
            Else
               vs.TextMatrix(J, 5) = -1
            End If
          
            con.Execute "update dnfa set Checked_YesNo=" & vs.TextMatrix(J, 5) & ",CheckedBy='" & UserName & "' where " & stringyear & " and dnn=" & vs.TextMatrix(J, 1) & ""
         
         
         
            If (AuditTrail = "y") Then
                    
                actionType_ = "Insert"
                vtype1_ = "D"
                vtypeNew = Left(vs.TextMatrix(J, 6), 1)
                vdate_ = vs.TextMatrix(J, 0)
                vno_ = vs.TextMatrix(J, 1)

                AuditTrail_Log "D", "Insert", "", vdate_, vno_, vtypeNew
        
            End If
         
         End If
         
        Next
   
   
End If
   
   

End If

search

Screen.MousePointer = vbDefault

Exit Sub
ss:
MsgBox err.Description, vbInformation
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()


cbovtype.Clear

If rs1.State = 1 Then rs1.Close
rs1.Open "SELECT VType_Det FROM CheckedBy where UName ='" & UserName & "'", con
While rs1.EOF = False

cbovtype.AddItem rs1(0)
rs1.MoveNext

Wend


Me.top = 100
Me.Left = 100

Me.Width = 18500
Me.Height = 11000

fDate_.value = Format(from_date, "dd/MM/yyyy")
tdate_.value = Format(to_date, "dd/MM/yyyy")

End Sub
Private Sub option_selectAll_Click()

For I = 1 To vs.rows - 1
   If (option_selectAll.value = 1) Then
      vs.TextMatrix(I, 5) = 1
   Else
      vs.TextMatrix(I, 5) = 0
   End If
Next

End Sub
