Attribute VB_Name = "mod_funcs"
Public Xtwips As Integer, Ytwips As Integer
Public Xpixels As Integer, Ypixels As Integer

'============================================
Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_WININICHANGE = &H1A
'============================================
Type FRMSIZE
  Height As Long
  Width As Long
End Type

'============================================
Public RePosForm As Boolean
Public DoResize As Boolean
 Public Function SetDefaultPrinter(objPrn As Printer) As Boolean
    Dim X As Long, sztemp As String
    sztemp = objPrn.DeviceName & "," & objPrn.DriverName & "," & objPrn.Port
    X = WriteProfileString("windows", "device", sztemp)
    X = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")
End Function
Sub Resize_For_Resolution(ByVal SFX As Single, _
ByVal SFY As Single, MyForm As Form)

Dim I As Integer
Dim SFFont As Single

SFFont = (SFX + SFY) / 2  ' average scale
' Size the Controls for the new resolution
On Error Resume Next  ' for read-only or nonexistent properties
With MyForm
  For I = 0 To .Count - 1
   If TypeOf .Controls(I) Is ComboBox Then   ' cannot change Height
     .Controls(I).Left = .Controls(I).Left * SFX
     .Controls(I).top = .Controls(I).top * SFY
     .Controls(I).Width = .Controls(I).Width * SFX
   Else
     .Controls(I).Move .Controls(I).Left * SFX, _
      .Controls(I).top * SFY, _
      .Controls(I).Width * SFX, _
      .Controls(I).Height * SFY
   End If
     ' Be sure to resize and reposition before changing the FontSize
     .Controls(I).FontSize = .Controls(I).FontSize * SFFont
  Next I
  
  If RePosForm Then
    'Now size the Form
    .Move .Left * SFX, .top * SFY, .Width * SFX, .Height * SFY
  End If

End With

End Sub

Sub close_previous_rs_form(mprevioustab As Integer)
    Select Case mprevioustab
    Case 0 'branches
        ' leave the RS open.
        Unload frmbranches
    Case 1 'users
        ' leave the RS open
        Unload frmusers
    Case 2 'vendors
        If mgot_through Then
            rs_vendors.Close
        End If
        Unload frmusers
    Case 3 'med_cats
        If mgot_through Then
            rs_med_cats.Close
        End If
        Unload frmmed_cats
    Case 4 'medicines
        If mgot_through Then
            rs_vendors.Close
            rs_med_cats.Close
            rs_medicines.Close
        End If
        Unload frmmedicines
    Case 5 'customers
        If mgot_through Then
            rs_customers.Close
            rs_tc_details.Close
            rs_tax_cats.Close
        End If
        Unload frmcustomers
    Case 6 'tax_cats
        If mgot_through Then
            rs_vendors.Close
            rs_tc_details.Close
            rs_tax_cats.Close
        End If
        Unload frmtax_cats
    Case 7 'purchases
        If mgot_through Then
            rs_med_cats.Close
            rs_medicines.Close
            rs_vendors.Close
            rs_tc_details.Close
            rs_tax_cats.Close
            rs_customers.Close
            rs_purchases2.Close
            rs_purchases1.Close
        End If
        
        Unload frmpurchases
    Case 8 'sales
        If mgot_through Then
            rs_med_cats.Close
            rs_medicines.Close
            rs_vendors.Close
            rs_tc_details.Close
            rs_tax_cats.Close
            rs_customers.Close
            rs_sales2.Close
            rs_sales1.Close
        End If
        Unload frmsales
    Case 11 'pay_rec
        If mgot_through Then
            rs_customers.Close
        End If
        Unload frmpay_rec
    Case 13 'cash sales
        Unload frmpcashs
    Case 14 'credit sales
        Unload frmpcredits
    Case 15 'party-wise sales
        If mgot_through Then
            rs_customers.Close
        End If
        Unload frmppws
    Case 16 'total sales
        Unload frmpts
    Case 17 'total sales/receipts
        If mgot_through Then
            rs_temp.Close
            rs_sales1.Close
            rs_sales2.Close
            rs_pay_rec.Close
        End If
        Unload frmptsr
    Case 18 'stock as on
        If mgot_through Then
            rs_temp.Close
            rs_medicines.Close
            rs_sales1.Close
            rs_sales2.Close
            rs_purchases1.Close
            rs_purchases2.Close
        End If
        Unload frmpbao
    Case 20 'balance as on
        If mgot_through Then
            rs_temp.Close
            rs_customers.Close
            rs_sales1.Close
            rs_sales2.Close
            rs_pay_rec.Close
        End If
        Unload frmpbao
    Case 20 'statement of a/c
        If mgot_through Then
            rs_temp.Close
            rs_customers.Close
            rs_sales1.Close
            rs_sales2.Close
            rs_pay_rec.Close
        End If
        Unload frmpstat
    Case 22 'company-wise stock as on
        If mgot_through Then
            rs_temp.Close
            rs_vendors.Close
            rs_medicines.Close
            rs_sales1.Close
            rs_sales2.Close
            rs_purchases1.Close
            rs_purchases2.Close
        End If
        Unload frmpvws
    End Select
End Sub
Sub HIT()
On Error Resume Next
  VB.Screen.ActiveForm.ActiveControl.SelStart = 0
  VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.text)
End Sub
Function search(msql As String, mlistfield As String, mreturnfield As String)
    search = ""
    Load frmSearch
    frmSearch.Data1.DatabaseName = mdatapath & "\smdmas.mdb"
    frmSearch.Data1.RecordSource = msql         '"select m_name, m_code from medicines order by m_name"
    frmSearch.DBCombo1.ListField = mlistfield   '"m_name"
    frmSearch.Data1.Refresh
    frmSearch.Show 1
    If Not Trim(frmSearch.DBCombo1.text) = "" Then
        frmSearch.Data1.Recordset.Bookmark = frmSearch.DBCombo1.SelectedItem
        search = frmSearch.Data1.Recordset.Fields(mreturnfield) 'Fields("m_code")
    End If
    Unload frmSearch
End Function

Function flock(mtable As String)
'On Error GoTo errortrap

    Dim mattempts As Integer
    flock = False
    rs_tables.Index = "Primary"
    rs_tables.MoveFirst
    rs_tables.Seek "=", mtable
    If rs_tables.NoMatch Then
        message "Internal Error. REQUIRED table is not existing"
        End
    End If
    mattempts = 0
    Do While mattempts < 500
        If rs_tables.Fields("locked") = False Then
            rs_tables.Edit
            If rs_tables.Fields("locked") = False Then
                rs_tables.Fields("locked") = True
                rs_tables.Fields("username") = mcurrentuser
                rs_tables.update
                flock = True
                Exit Do
            Else
                rs_tables.CancelUpdate
            End If
        End If
        mattempts = mattempts + 1
    Loop
    
End Function
Function funlock(mtable As String)
'On Error GoTo errortrap

    funlock = False
    rs_tables.Index = "Primary"
    rs_tables.Seek "=", mtable
    If rs_tables.NoMatch Then
        message "Internal Error. REQUIRED table is not existing"
        End
    End If
    rs_tables.Edit
    rs_tables.Fields("locked") = False
    rs_tables.Fields("username") = ""
    rs_tables.update
    funlock = True
   
'errortrap:
 '   errorcall

End Function

Function mystr(mvar As String, MLENGTH As Integer, MDECIMAL As Boolean)
    Dim mtemp As String
    mtemp = String(MLENGTH, "0")
    If Not MDECIMAL Then    'InStr(1, mvar, ".", 1) = 0 Then
        mvar = Format(Val(Trim(mvar)), "0")
    Else
        If Val(mvar) <> 0 Then
        If Int(mvar) = mvar Then
            mvar = Trim(mvar) & ".00"
        End If
                mvar = Format(Val(Trim(mvar)), numformat)
    End If
        End If
    RSet mtemp = mvar
    mystr = mtemp
   End Function

Function oper_allowed(mvar1 As Integer, mvar2 As Integer)
    oper_allowed = False
    rs_urights.Index = "Primary"
    rs_urights.Seek "=", mcurrentuser, mvar1
    If Not rs_urights.NoMatch Then
        Select Case mvar2
            Case allow_add
                If rs_urights.Fields("addmode") = True Then
                    oper_allowed = True
                End If
            Case allow_edit
                If rs_urights.Fields("editmode") = True Then
                    oper_allowed = True
                End If
            Case allow_delete
                If rs_urights.Fields("deletemode") = True Then
                    oper_allowed = True
                End If
            Case allow_view
                If rs_urights.Fields("viewmode") = True Then
                    oper_allowed = True
                End If
        End Select
    End If
    If oper_allowed = False Then
        message "Not allowed !"
        mgot_through = False
    Else
        mgot_through = True
    End If
End Function
Sub message(mvar As String)
    MsgBox mvar, , " ERROR !"
End Sub
Sub quit()
End Sub


Sub open_table(mrs As Recordset, mtable As String)
    Set mrs = db_bdpl.OpenRecordset(mtable)
End Sub
Function valid_date(mvar As String)
    If Len(Trim(mvar)) < 8 Or Val(Mid(mvar, 3, 2)) > 12 Then
        valid_date = ""
        Exit Function
    End If
    mvar = Left(mvar, 2) & "/" & Mid(mvar, 3, 2) & "/" & Right(mvar, 4)
    valid_date = mvar
End Function
Sub SetButton(ParamArray Bt())

Dim rsuser As New ADODB.Recordset


If rs1.State = 1 Then rs1.Close
If mnuMenu_ <> "" Then
    rs1.Open "Select * from UsrePermission where TaskType = '" & mnuMenu_ & "' and username = '" & UserName & "' and formWiseP='y'  order by userid", coninfo, adOpenKeyset, adLockOptimistic
    If rs1.EOF = True Then
       mnuMenu_ = ""
    End If
End If


If rsuser.State = 1 Then rsuser.Close
If mnuMenu_ = "" Then
   rsuser.Open "Select top 2  * from UsrePermission where [Module]='" & module_ & "' and username = '" & UserName & "'  order by userid", coninfo, adOpenKeyset, adLockOptimistic
Else
   rsuser.Open "Select * from UsrePermission where TaskType = '" & mnuMenu_ & "' and username = '" & UserName & "'", coninfo, adOpenKeyset, adLockOptimistic
End If
   
   If rsuser.RecordCount > 0 Then
      c = 0
      For Each I In Bt
         c = c + 1

     If c = 1 Then
        If rsuser!save = "y" Then
             Bt(0).Enabled = True
        Else
             Bt(0).Enabled = False
        End If
     End If
      
     If c = 2 Then
        If rsuser!Edit = "y" Then
             Bt(1).Enabled = True
            
        Else
             Bt(1).Enabled = False
             
        End If
     End If
     If c = 4 Then
         If rsuser!delete = "y" Then
            Bt(3).Enabled = True
        Else
            Bt(3).Enabled = False
      End If
    End If
    If c = 3 Then
        If rsuser!save = "y" Then
           Bt(2).Enabled = True
        Else
           Bt(2).Enabled = False
        End If
    End If
      
   Next I
   End If
   
   mnuMenu_ = ""
   
End Sub
Sub fillcombo(c As Control, Field, table, adod As ADODB.Connection, Optional fld As String, Optional val1 As String)
    
    On Error Resume Next
    If CStr(Field) = "" Then Exit Sub
    If fld = "" Then
    Set RS = adod.Execute("select " & Field & " from " & table & "  where " & stringyear & " group by " & Field & " order by " & Field)
    Else
    Set RS = adod.Execute("select " & Field & " from " & table & " where " & stringyear & " and cstr(" & fld & ")='" & val1 & "' group by " & Field & " order by " & Field)
    End If
    
    If Not RS.EOF Then
        c.Clear
        While Not RS.EOF
        If Not IsNull(RS(0)) Then
            c.AddItem Trim(UCase(RS(0)))
        End If
            RS.MoveNext
        Wend
        c.ListIndex = 0
    End If
    
End Sub


