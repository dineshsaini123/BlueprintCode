VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPaperLedger 
   Caption         =   "Paper Ledger ..."
   ClientHeight    =   7704
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   12420
   Icon            =   "frmPaperLedger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7704
   ScaleWidth      =   12420
   Begin VB.TextBox txtFrom_Gdid 
      Height          =   285
      Left            =   6570
      MaxLength       =   50
      TabIndex        =   23
      Top             =   990
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton cmdOp 
      BackColor       =   &H00BFFFBF&
      Caption         =   "&Closing Transfer"
      Height          =   480
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1485
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPcode 
      Enabled         =   0   'False
      Height          =   345
      Left            =   6540
      TabIndex        =   19
      Top             =   120
      Width           =   450
   End
   Begin Crystal.CrystalReport cr 
      Left            =   120
      Top             =   7200
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   480
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1485
      Width           =   1245
   End
   Begin VB.TextBox txtSheet_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10215
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1665
      Width           =   915
   End
   Begin VB.TextBox txtReam_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9180
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtSheet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10140
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6945
      Width           =   915
   End
   Begin VB.TextBox txtReams 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9165
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6945
      Width           =   930
   End
   Begin VB.ComboBox cboGodown 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   10
      Top             =   960
      Width           =   4815
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   480
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1485
      Width           =   1305
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   480
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1485
      Width           =   1245
   End
   Begin VB.ComboBox cboPaperSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   1800
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   120
      Width           =   4755
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4545
      Left            =   60
      TabIndex        =   1
      Top             =   2160
      Width           =   11265
      _cx             =   19870
      _cy             =   8017
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12582847
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   450
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPaperLedger.frx":000C
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
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
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   13440
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   3960
         Width           =   195
      End
   End
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   540
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   656
      _Version        =   393216
      Format          =   184090625
      CurrentDate     =   38531
   End
   Begin MSComCtl2.DTPicker toDate 
      Height          =   375
      Left            =   3450
      TabIndex        =   7
      Top             =   525
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   656
      _Version        =   393216
      Format          =   184090625
      CurrentDate     =   38531
   End
   Begin VB.Label lblP_det 
      Height          =   315
      Left            =   11340
      TabIndex        =   21
      Top             =   1860
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblPaper_det 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   7080
      TabIndex        =   20
      Top             =   120
      Width           =   4155
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      FillColor       =   &H00C0C0FF&
      Height          =   435
      Left            =   7530
      Top             =   6885
      Width           =   3690
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      FillColor       =   &H00C0C0FF&
      Height          =   435
      Left            =   7500
      Top             =   1620
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Balance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7650
      TabIndex        =   14
      Top             =   7005
      Width           =   1515
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer/Godown :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   1020
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   615
      Width           =   1305
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3105
      TabIndex        =   8
      Top             =   555
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Size :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   180
      Width           =   1515
   End
End
Attribute VB_Name = "frmPaperLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ream_tot, sheet_tot
Dim dateRange
Dim pname As String
Private Sub cboGodown_Click()
txtFrom_Gdid = cboGodown.ItemData(cboGodown.ListIndex)
End Sub

Private Sub cboPaperSize_GotFocus()

If PopUpValue1 <> "" Then
   
   
   txtPcode.text = popupvalue5
   
   
   If Right(Trim(PopUpValue3), 1) = "X" Then
      PopUpValue3 = Trim(PopUpValue3)
      PopUpValue3 = Mid(PopUpValue3, 1, Len(PopUpValue3) - 1) & "CM"
      cboPaperSize.text = PopUpValue3
   End If
   
   lblPaper_det.Caption = "Paper Name : " & PopUpValue1 & vbCrLf & "Paper Type : " & PopUpValue2 & vbCrLf & "Size : " & PopUpValue3 & vbCrLf & "G.S.M. : " & popupvalue4
   lblP_det.Caption = popupvalue5 & "   GSM"
       
   If RS.State = 1 Then RS.close
   RS.Open "select * from PaperMakeMaster where papermaker_id='" & txtPcode.text & "'", con, adOpenStatic, adLockReadOnly
   If RS.EOF = False Then
      
      pname = RS!papermaker_name
      If RS!eco <> "" Then
       pname = pname & "-" & RS!eco
      End If
    
      If RS!SizeValue1 <> "" Then
         pname = pname & "-" & RS!SizeValue1 & "X" & RS!SizeValue2
      End If
    
      If RS!GSM <> "" Then
         pname = pname & "-" & RS!GSM
      End If
      
    End If
   
   cboPaperSize.text = pname
    
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""

End If
  

End Sub

Private Sub cboPaperSize_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
       
'   searchType = "paper"
'   value = "Select SizeValue1 + 'X'+ SizeValue2 as [Paper Size]," & _
'   "papermaker_name as [Paper Name],PType,Size as [Sheet/Real],Eco + '   -  ' + GSM  as [Quality & GSM],papermaker_Id as Code from " & _
'   " papermakemaster where " & stringyear & " and papermaker_id <> ''"
'    popuplistModel10 value, con
    
    
    
   value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear & " order by papermaker_name"
   popuplist1 value, con
 
    

End If

End Sub

Private Sub cmdexit_Click()

'If RS.State = 1 Then RS.close
'RS.Open "select * from PaperMakeMaster order by papermaker_id", con, adOpenStatic, adLockReadOnly
'
'While RS.EOF = False
'    pname = ""
'
'    If RS!papermaker_id = 104 Then
'       MsgBox ""
'    End If
'
'    pname = RS!papermaker_name
'
'    If Len(RS!eco) > 1 Then
'    If RS!eco <> "" Then
'       pname = pname & "-" & RS!eco
'    End If
'    End If
'
'    If RS!SizeValue1 <> "" Then
'       pname = pname & "-" & RS!SizeValue1
'    End If
'
'    If RS!SizeValue2 <> "" Then
'       pname = pname & "X" & RS!SizeValue2
'    End If
'
'
'    If RS!GSM <> "" Then
'       pname = pname & "-" & RS!GSM
'    End If
'
'    con.Execute "update PaperMakeMaster set papername1='" & pname & "' where papermaker_id='" & RS!papermaker_id & "'"
'
'RS.MoveNext
'Wend
'



'''''
Unload Me
End Sub
Sub vs_ini()

  vs.Clear
  vs.Cols = 10
  'vs.FormatString = "Date|Particulars|Order No|Challan No|Received|Devivered/Consumed|Balance"
  
  vs.TextMatrix(0, 0) = "Date"
  vs.TextMatrix(1, 0) = "Date"

  vs.TextMatrix(0, 1) = "Particulars"
  vs.TextMatrix(1, 1) = "Particulars"
  
  vs.TextMatrix(0, 2) = "Order No"
  vs.TextMatrix(1, 2) = "Order No"
  
  vs.TextMatrix(0, 3) = "Challan No"
  vs.TextMatrix(1, 3) = "Challan No"
  
  vs.TextMatrix(0, 4) = "Received"
  vs.TextMatrix(0, 5) = "Received"
  
  vs.TextMatrix(1, 4) = "Reams"
  vs.TextMatrix(1, 5) = "Sheets"
  
  vs.TextMatrix(0, 6) = "Del/Consumed"
  vs.TextMatrix(0, 7) = "Del/Consumed"
  
  vs.TextMatrix(1, 6) = "Reams"
  vs.TextMatrix(1, 7) = "Sheets"
  
  
  vs.TextMatrix(0, 8) = "Balance"
  vs.TextMatrix(0, 9) = "Balance"
  
  vs.TextMatrix(1, 8) = "Reams"
  vs.TextMatrix(1, 9) = "Sheets"
  
  
  vs.MergeCells = flexMergeFixedOnly

For I = 0 To 8
   vs.MergeCol(I) = True
   'vs.MergeRow(I) = True
   ''vs.Cell(flexcpFontSize, 0, I) = 10
   ''vs.Cell(flexcpFontSize, 1, I) = 8
Next

vs.WordWrap = True
  
vs.ColWidth(0) = 1700
vs.ColWidth(1) = 4200
vs.ColWidth(2) = 1300
vs.ColWidth(3) = 1500
vs.ColWidth(4) = 1200
vs.ColWidth(5) = 1200
vs.ColWidth(6) = 1450
vs.ColWidth(7) = 1450
vs.ColWidth(8) = 1780
vs.ColWidth(9) = 1780
  
  
  
  

End Sub

Private Sub cmdOp_Click()


   Dim CON_next As New ADODB.Connection
   Dim db_ As String

   Dim a1 As Integer
   

   
   a1 = Int(Right(databaseNew, 4)) + 101
   
   db_ = "Database=chitraData_" & a1
   Set CON_next = New ADODB.Connection
   CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
    
   DoEvents
   DoEvents
    
    CON_next.CursorLocation = adUseClient
    If CON_next.State = 1 Then CON_next.close
    CON_next.Open

  
   
     
'Set RS = New ADODB.Recordset
'RS.Open "paperstatement where (sn_no = 0  and PaperTrans_Deliv = 'R' and FromGodown='" & cboGodown.text & "')", CON_next, adOpenDynamic, adLockPessimistic
'If RS.EOF = False Then

'   If MsgBox("Closing Already Transfer , Are you sure want to transfer Again... ", vbCritical + vbYesNo) = vbNo Then
'      Exit Sub
'   End If
   
'End If


Dim dd_ As String
dd_ = ""
dd_ = DateAdd("d", 1, toDate)



'sq = "delete from paperstatement where (sn_no = 0  and PaperTrans_Deliv = 'R' and FromGodown='" & cboGodown.text & "')"
'CON_next.Execute sq


Set RS = New ADODB.Recordset
RS.Open "paperstatement", CON_next, adOpenDynamic, adLockPessimistic
For I = 1 To vs.rows - 1

 If vs.TextMatrix(I, 0) <> "" Then

         sq = "delete from paperstatement where (Challan_No = 'Op.Bal.' and pcode='" & vs.TextMatrix(I, 4) & "' and FromGodown='" & cboGodown.text & "')"
        CON_next.Execute sq


        RS.AddNew
        RS.Fields("PaperTrans_Deliv").value = "R"
        RS.Fields("Challan_No") = "Op.Bal."
        RS.Fields("Challan_Date") = dd_
        RS.Fields("Sn_No") = 0
        RS.Fields("Sn_Date") = dd_
        RS.Fields("remarks") = "Opening"
        
        RS.Fields("FromGodown") = cboGodown.text
        RS.Fields("ToGodown") = ""
        
        RS.Fields("FromGodown_id") = txtFrom_Gdid.text
        RS.Fields("ToGodown_id") = ""
        
        
        RS.Fields("sno") = vs.TextMatrix(I, 0)
        
        RS.Fields("Bill_Date") = "0"
        
        
        If rs1.State = 1 Then rs1.close
        rs1.Open "select SizeValue1,SizeValue2,gsm,PaperName1 from PaperMakeMaster where papermaker_id='" & vs.TextMatrix(I, 4) & "'", con
        If rs1.EOF = False Then
           RS.Fields("size") = rs1(0) & "X" & rs1(1)
           RS.Fields("gsm") = rs1(2)
           RS.Fields("pcode") = vs.TextMatrix(I, 4)
           RS.Fields("Paper_Make") = rs1!PaperName1 & "=>" & vs.TextMatrix(I, 4)
        End If
        
        
        RS.Fields("reams") = IIf(vs.TextMatrix(I, 2) = "", 0, vs.TextMatrix(I, 2))
        RS.Fields("Sheets") = Val(vs.TextMatrix(I, 3))
        v1 = Left(session, 4) + 1
        v2 = Right(session, 2) + 1
        
        RS.Fields("fyear") = v1 & "-" & v2     'session
        RS.Fields("setupid") = setupid
        
        RS.update
 
 End If
 
 Next
   
MsgBox "Transfer....", vbInformation






End Sub

Private Sub cmdPrint_Click()

If vs.Cols <= 5 Then
   'fillExcel
   If RS.State = 1 Then RS.close
   con.Execute "delete from TmpBook1"
   For I = 1 To vs.rows - 1
     If vs.TextMatrix(I, 1) <> "" Then
        con.Execute "insert into TmpBook1(Qty,BName,states,city) values('" & vs.TextMatrix(I, 0) & "','" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & Format(Abs(vs.TextMatrix(I, 3)), "000") & "')"
     End If
   Next
   
   Screen.MousePointer = vbHourglass
   
    DSNNew
    
    cr.Reset
    cr.ReportFileName = rptPath & "/PaperSt_PrinterWise.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    
    If cboGodown.text <> "" Then
       cr.Formulas(0) = "printerName='" & cboGodown.text & "'"
       cr.Formulas(1) = "pName='" & "Paper Name" & "'"
    Else
       cr.Formulas(0) = "printerName='" & pname & "'"
       cr.Formulas(1) = "pName='" & "Printer Name" & "'"
    End If
    
    
    
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowRefreshBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1

   Screen.MousePointer = vbDefault
   
   Exit Sub
End If


If RS.State = 1 Then RS.close
RS.Open "select * from PrintPaperLedger", con, adOpenDynamic, adLockOptimistic

For a = 2 To vs.rows - 1

If vs.TextMatrix(a, 0) <> "" Then

    RS.AddNew
    RS!Size = Trim(cboPaperSize.text) '& "  :  " & lblP_det
    RS!FromDate = FromDate.value
    RS!toDate = toDate.value
    If cboGodown.text <> "" Then
    RS!Printer_Godown = cboGodown.text
    End If
    RS!Reams_Op = Val(txtReam_op.text)
    RS!Sheets_Op = Val(txtSheet_op.text)
    
    RS!dates = vs.TextMatrix(a, 0)
        
    RS!Particulars = vs.TextMatrix(a, 1)
    
    RS!orderNo = IIf(vs.TextMatrix(a, 2) = "", "-", vs.TextMatrix(a, 2))
    RS!ChallanNo = IIf(vs.TextMatrix(a, 3) = "", "-", vs.TextMatrix(a, 3))
    RS!Reams_Rec = Val(vs.TextMatrix(a, 4))
    RS!Sheets_Rec = Val(vs.TextMatrix(a, 5))
    RS!Reams_Del = Val(vs.TextMatrix(a, 6))
    RS!Sheets_Del = Val(vs.TextMatrix(a, 7))
    
    RS!Reams_Bal = vs.TextMatrix(a, 8)
    RS!Sheets_Bal = vs.TextMatrix(a, 9)
    
    RS!Reams_Clos = Val(txtReams.text)
    RS!Sheets_Clos = Val(txtSheet.text)
    
    RS.update
    
End If

Next


'---------------------------------------------
DSNNew

cr.Reset
cr.ReportFileName = rptPath & "/PaperLedger.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.WindowShowPrintSetupBtn = True
cr.WindowState = crptMaximized
cr.Action = 1

lblP_det = ""

'---------------------------------------------
cmdprint.Enabled = False

End Sub
Public Function Cal_ReamAndSheet(ByVal rm_ As Long, ByVal st_ As Long)
    
Dim cal_ream
Dim D As Integer
Dim recSheet, tmpReam, tmpSheet
recSheet = 0
tmpReam = 0
tmpSheet = 0


recSheet = (rm_ * 500) + st_

If recSheet < 0 Then

recSheet = Abs(recSheet)

tmpReam = Int(recSheet / 500)
tmpSheet = (Round(((recSheet / 500) - Int(recSheet / 500)), 3) * 1000 / 2)
ream_tot = (tmpReam * -1)
sheet_tot = (tmpSheet * -1)

Else

tmpReam = Int(recSheet / 500)
tmpSheet = (Round(((recSheet / 500) - Int(recSheet / 500)), 3) * 1000 / 2)
ream_tot = tmpReam
sheet_tot = tmpSheet


End If

End Function
Sub fillTotalPaper_Printerwise()


Dim rs_ As New ADODB.Recordset
Dim rs_1 As New ADODB.Recordset

vs.Clear
vs.Cols = 5
vs.rows = 2

stDate = "(convert(smalldatetime,Dates,103)>= convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & toDate.value & "',103))"
If cboGodown.text = "DHARMESH ART PROCESS & PRINT (P) LTD." Then
If session = "2019-20" Then
   stDate = "(convert(smalldatetime,Dates,103)>= convert(smalldatetime,'" & "30/05/2019" & "',103) and convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & toDate.value & "',103))"
End If
End If


s = ""


If cboGodown.text <> "" Then

    If rs_.State = 1 Then rs_.close
    rs_.Open "SELECT Pcode,sum(Sheet) FROM TotalpaperReamsPrinterWise where printer='" & cboGodown.text & "' and " & stDate & " group by Pcode", con
    For J = 1 To rs_.RecordCount
    
    If rs_1.State = 1 Then rs_1.close
    rs_1.Open "select * from PaperMakeMaster where papermaker_id='" & rs_!pcode & "'", con, adOpenStatic, adLockReadOnly
    If rs_1.EOF = False Then
         s = rs_1!papermaker_name
         If rs_1!eco <> "" Then s = s & "-" & rs_1!eco
         If rs_1!SizeValue1 <> "" Then s = s & "-" & rs_1!SizeValue1 & "X" & rs_1!SizeValue2
         If rs_1!GSM <> "" Then s = s & "-" & rs_1!GSM
    
    
    End If
    
    vs.TextMatrix(J, 0) = J
    vs.TextMatrix(J, 1) = rs_1!PaperName1 & ""
    Cal_ReamAndSheet 0, rs_(1)
    vs.TextMatrix(J, 2) = ream_tot
    vs.TextMatrix(J, 3) = sheet_tot
    vs.TextMatrix(J, 4) = rs_!pcode
    
    vs.rows = vs.rows + 1
    rs_.MoveNext
    
    Next
    
    vs.FormatString = "SN|Paper Size|>Total Reams|>Total Sheet"
    vs.ColWidth(0) = 1000
    vs.ColWidth(1) = 8500
    vs.ColWidth(2) = 1800
    vs.ColWidth(3) = 1800
    vs.ColWidth(4) = 0

Else

J = 1

For kk1 = 0 To cboGodown.ListCount - 1

    If rs_.State = 1 Then rs_.close
    rs_.Open "SELECT pcode,sum(Sheet),Printer FROM TotalpaperReamsPrinterWise where (printer='" & cboGodown.List(kk1) & "' and Pcode='" & txtPcode.text & "' and " & stDate & ") group by Printer,pcode", con
    If rs_.EOF = False Then
    'For J = 1 To rs_.RecordCount
    
    If rs_1.State = 1 Then rs_1.close
    rs_1.Open "select * from PaperMakeMaster where papermaker_id='" & rs_!pcode & "'", con, adOpenStatic, adLockReadOnly
    If rs_1.EOF = False Then
         s = rs_1!papermaker_name
         If rs_1!eco <> "" Then s = s & "-" & rs_1!eco
         If rs_1!SizeValue1 <> "" Then s = s & "-" & rs_1!SizeValue1 & "X" & rs_1!SizeValue2
         If rs_1!GSM <> "" Then s = s & "-" & rs_1!GSM
    End If
    
    vs.TextMatrix(J, 0) = J
    vs.TextMatrix(J, 1) = rs_!Printer
    Cal_ReamAndSheet 0, rs_(1)
    vs.TextMatrix(J, 2) = ream_tot
    vs.TextMatrix(J, 3) = sheet_tot
    vs.TextMatrix(J, 4) = rs_!pcode
    
    vs.rows = vs.rows + 1
    J = J + 1
    
    rs_.MoveNext
    
    'Next
    End If
    
 Next
    
    vs.FormatString = "SN|Printer Name|>Total Reams|>Total Sheet"
    vs.ColWidth(0) = 1000
    vs.ColWidth(1) = 6500
    vs.ColWidth(2) = 1800
    vs.ColWidth(3) = 1800
    vs.ColWidth(4) = 0


End If


End Sub

Sub fillExcel()

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
Dim Q1, q2, J As Double

Dim b1 As Boolean

b1 = False


c = 1
r = 1





    row_ = 1
    col_ = 1
   
    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    'xlSheet.Cells(1, 1).value = "Book Wise Ordered Qty "
    
    For I = 0 To vs.rows - 1
        For J = 0 To vs.Cols - 1
               xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
              col_ = col_ + 1
        Next
        row_ = row_ + 1
        col_ = 1
    Next
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
err:
    
    


End Sub
Private Sub Commandshow_Click()

Screen.MousePointer = vbHourglass

Dim rm, ST, vs_raw As Integer



rm = 0
ST = 0
txtReams = 0
txtSheet = 0
txtReam_op = 0
txtSheet_op = 0

vs_raw = 2
vs.FixedRows = 2

If (cboPaperSize.text = "" Or cboGodown.text = "") Then

   vs.FixedRows = 1
   fillTotalPaper_Printerwise
   cmdprint.Enabled = True
   Screen.MousePointer = vbDefault
   Exit Sub
Else
   vs_ini
End If


dateRange = "convert(smalldatetime,Ord_Date,103)< convert(smalldatetime,'" & FromDate.value & "',103)"

'If cboGodown.Text = "DHARMESH ART PROCESS & PRINT (P) LTD." Then
'If session = "2019-20" Then
'   dateRange = "not convert(smalldatetime,Ord_Date,103) between convert(smalldatetime,'" & "02/04/2019" & "',103) and convert(smalldatetime,'" & "30/05/2019" & "',103)"
'End If
'End If

ClosingBal

con.Execute "delete from TMPpaperstatement where (username='" & UserName & "')"


vs.rows = 100
If RS.State = 1 Then RS.close



   If cboGodown.text = "" Then
        RS.Open "select Ord_Date,PrinterName,Ord_No,Ream,Sheet from Order_Qry " & _
        " where ((convert(smalldatetime,Ord_Date,103)>= convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,Ord_Date,103)<= convert(smalldatetime,'" & toDate.value & "',103)) and " & _
        " Pcode='" & txtPcode & "') order by Ord_No,Ord_Date", con, adOpenKeyset, adLockReadOnly
    Else
       
        RS.Open "select Ord_Date,PrinterName,Ord_No,Ream,Sheet from Order_Qry " & _
        " where ((convert(smalldatetime,Ord_Date,103)>= convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,Ord_Date,103)<= convert(smalldatetime,'" & toDate.value & "',103)) and " & _
        " PrinterName='" & cboGodown.text & "' and Pcode='" & txtPcode & "') order by Ord_No,Ord_Date", con, adOpenKeyset, adLockReadOnly
        
    End If




For k1 = 2 To RS.RecordCount + 2
 
    If RS.EOF = False Then
        
        
        Cal_ReamAndSheet RS!ream, RS!sheet
        
        
        rm = rm - ream_tot
        ST = ST - sheet_tot
        
        
        con.Execute "insert into TMPpaperstatement(Dates,Narr,OrderNo,Rec_Reams,Rec_Sheets,Del_Reams,Del_Sheets,Bal_Reams,Printer,username) " & _
        " values('" & Format(RS!Ord_Date, "MM/dd/yyyy") & "','" & RS!PrinterName & "','" & RS!Ord_No & "',0,0," & ream_tot & "," & sheet_tot & ",0,'" & RS!PrinterName & "','" & UserName & "')"
        
        vs_raw = vs_raw + 1
        RS.MoveNext
    End If
Next

'------------------------------Data Fatch Paper Rec & Deliver-----------------------------------------

If RS.State = 1 Then RS.close

If cboGodown.text = "" Then

    RS.Open "select Challan_date,FromGodown,Challan_No,reams,sheets,PaperTrans_Deliv,toGodown,SN_Date from paperstatement " & _
    " where (convert(smalldatetime,Challan_date,103)>= convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,Challan_date,103)<= convert(smalldatetime,'" & toDate.value & "',103)) and " & _
    " PaperTrans_Deliv='R' and pcode='" & txtPcode & "' order by sn_date,sn_no", con, adOpenKeyset, adLockReadOnly

Else

    RS.Open "select Challan_date,FromGodown,Challan_No,reams,sheets,PaperTrans_Deliv,toGodown,SN_Date from paperstatement " & _
    " where (convert(smalldatetime,SN_Date,103)>= convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,SN_Date,103)<= convert(smalldatetime,'" & toDate.value & "',103)) and " & _
    " PaperTrans_Deliv='R' and (FromGodown='" & cboGodown.text & "' or toGodown='" & cboGodown.text & "') and pcode='" & txtPcode & "' order by sn_date,sn_no", con, adOpenKeyset, adLockReadOnly

End If

For k1 = 2 To RS.RecordCount + 2
    If RS.EOF = False Then
        

        con.Execute "insert into TMPpaperstatement(Dates,Narr,challanNo,Rec_Reams,Rec_Sheets,Del_Reams,Del_Sheets,Bal_Reams,Printer,username) " & _
        " values('" & Format(RS!SN_Date, "MM/dd/yyyy") & "','" & RS!FromGodown & "','" & RS!Challan_No & "'," & RS!reams & "," & RS!Sheets & ",0,0,0,'" & RS!FromGodown & "','" & UserName & "')"

        
        rm = rm + RS!reams
        ST = ST + RS!Sheets
        
        
    
    RS.MoveNext
    End If

Next



'------------------------------Transfer Rec-----------------------------------------

If RS.State = 1 Then RS.close

If cboGodown.text = "" Then

    RS.Open "select Challan_date,FromGodown,Challan_No,reams,sheets,PaperTrans_Deliv,toGodown,SN_Date from paperstatement " & _
    " where (convert(smalldatetime,Challan_date,103)>= convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,Challan_date,103)<= convert(smalldatetime,'" & toDate.value & "',103)) and " & _
    " (PaperTrans_Deliv='D') and pcode='" & txtPcode & "' order by sn_date,sn_no", con, adOpenKeyset, adLockReadOnly

Else

    RS.Open "select Challan_date,FromGodown,Challan_No,reams,sheets,PaperTrans_Deliv,toGodown,SN_Date from paperstatement " & _
    " where (convert(smalldatetime,SN_Date,103)>= convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,SN_Date,103)<= convert(smalldatetime,'" & toDate.value & "',103)) and " & _
    " (PaperTrans_Deliv='D') and (toGodown='" & cboGodown.text & "') and pcode='" & txtPcode & "' order by sn_date,sn_no", con, adOpenKeyset, adLockReadOnly

End If

For k1 = 2 To RS.RecordCount + 2
    If RS.EOF = False Then
        

        con.Execute "insert into TMPpaperstatement(Dates,Narr,challanNo,Rec_Reams,Rec_Sheets,Del_Reams,Del_Sheets,Bal_Reams,Printer,username) " & _
        " values('" & Format(RS!SN_Date, "MM/dd/yyyy") & "','" & cboGodown.text & "','" & RS!Challan_No & "'," & RS!reams & "," & RS!Sheets & ",0,0,0,'" & cboGodown.text & "','" & UserName & "')"
       
        
        rm = rm + RS!reams
        ST = ST + RS!Sheets
        
        
    
    RS.MoveNext
    End If

Next







'------------------------------Paper Delivered---------------------------------------------------------------------------
If RS.State = 1 Then RS.close

If cboGodown.text = "" Then
    RS.Open "select SN_date,FromGodown as PrinterName,Challan_No,reams,sheets,ToGodown from paperstatement " & _
    " where (convert(smalldatetime,SN_date,103)>= convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,SN_date,103)<= convert(smalldatetime,'" & toDate.value & "',103)) and " & _
    " PaperTrans_Deliv='D' order by sn_date,sn_no", con, adOpenKeyset, adLockReadOnly
Else
    RS.Open "select SN_date,FromGodown as PrinterName,Challan_No,reams,sheets,ToGodown from paperstatement " & _
    " where (convert(smalldatetime,SN_date,103)>= convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,SN_date,103)<= convert(smalldatetime,'" & toDate.value & "',103)) and " & _
    " PaperTrans_Deliv='D' and  FromGodown='" & cboGodown.text & "' and pcode='" & txtPcode & "' and  " & stringyear & " order by sn_date,sn_no", con, adOpenKeyset, adLockReadOnly
End If

For k1 = 2 To RS.RecordCount + 2
    If RS.EOF = False Then
        
        vs.TextMatrix(vs_raw, 0) = RS!SN_Date
        
       If cboGodown.text = "" Then
          vs.TextMatrix(vs_raw, 1) = RS!FromGodown & " to " & RS!toGodown
       Else
          vs.TextMatrix(vs_raw, 1) = UCase("Transfar to ") & RS!toGodown
       End If
        
        
        
        vs.TextMatrix(vs_raw, 3) = RS!Challan_No
        vs.TextMatrix(vs_raw, 6) = RS!reams
        vs.TextMatrix(vs_raw, 7) = RS!Sheets
        vs.TextMatrix(vs_raw, 8) = RS!reams
        vs.TextMatrix(vs_raw, 9) = RS!Sheets
        
        
        vs.Cell(flexcpFontSize, vs_raw, 1) = 7.5
        rm = rm - RS!reams
        ST = ST - RS!Sheets
        
        con.Execute "insert into TMPpaperstatement(Dates,Narr,challanNo,Rec_Reams,Rec_Sheets,Del_Reams,Del_Sheets,Bal_Reams,Printer,userName) " & _
        " values('" & Format(RS!SN_Date, "MM/dd/yyyy") & "','" & RS!PrinterName & "','" & RS!Challan_No & "',0,0," & RS!reams & "," & RS!Sheets & ",0,'" & RS!PrinterName & "','" & UserName & "')"

    
    RS.MoveNext
    End If

Next



rm = rm + Val(txtReam_op)
ST = ST + Val(txtSheet_op)

Cal_ReamAndSheet rm, ST

txtReams.text = ream_tot
txtSheet.text = sheet_tot

con.Execute "delete from PrintPaperLedger where (len(Size)>0)"




If (cboGodown.text = "" And cboPaperSize.text <> "") Then
    vs.Cols = 4
    vs.Clear
    
    vs.FormatString = "SN|Paper Size|>Total Reams|>Total Sheet"
    vs.ColWidth(0) = 1000
    vs.ColWidth(1) = 8000
    vs.ColWidth(2) = 2000
    vs.ColWidth(3) = 2000
Else
    vs_ini
End If


'---------------------------------------------------------
Dim rm1, op_
ream_tot = 0
sheet_tot = 0
    

op_ = IIf(txtReam_op = "", 0, txtReam_op) * 500
op_ = op_ + IIf(txtSheet_op = "", 0, txtSheet_op)


''If rs1.State = 1 Then rs1.close
''rs1.Open "select Rec_Reams,Rec_Sheets,Del_Reams,Del_Sheets,aouto from TMPpaperstatement where (Printer='" & cboGodown.Text & "' and username='" & UserName & "') order by dates"
''
''If rs1.EOF = False Then
''  rm1 = op_ + (rs1!Rec_Reams * 500) + rs1!Rec_Sheets
''End If
''
''While rs1.EOF = False
''
''    rm1 = (rs1!Rec_Reams * 500) + rs1!Rec_Sheets + op_
''    rm1 = rm1 + ((rs1!Del_Reams * 500) * -1 + (rs1!Del_Sheets * -1))
''
''    If RS.State = 1 Then RS.close
''    RS.Open "select Bal_Reams,Bal_Sheets,aouto from TMPpaperstatement where (aouto='" & rs1!aouto & "' and Printer='" & cboGodown.Text & "' and username='" & UserName & "')", con
''    If RS.BOF = False Then
''       rm1 = rm1 + (RS!Bal_Reams * 500) + RS!Bal_Sheets
''    End If
''
''    If Not IsNull(rm1) Then
''       Cal_ReamAndSheet 0, rm1
''    End If
''
''    con.Execute "update TMPpaperstatement set Bal_Reams=" & ream_tot & ",Bal_Sheets=" & sheet_tot & " where aouto=" & rs1!aouto & ""
''
''    op_ = 0
''
''rs1.MoveNext
''Wend


'Fill Grid------------------------------------------------

rm2 = 0

If (cboGodown.text = "" And cboPaperSize.text <> "") Then
    
    vs.FixedRows = 1
    
    If RS.State = 1 Then RS.close
    RS.Open "SELECT [Narr],sum(Rec_Reams) as Rreams,sum(Rec_Sheets) as RSheet,sum(Del_Reams) as DReams,sum(Del_Sheets) as DSheet from TMPpaperstatement where username='" & UserName & "' group by Narr order by date", con
    
    For J = 1 To RS.RecordCount
      
      rm1 = (RS(1) * 500) + RS(2)
      rm2 = (RS(3) * 500) + RS(4)
      
      rm1 = rm1 - rm2
      
      If Not IsNull(rm1) Then
         Cal_ReamAndSheet 0, rm1
      End If
    
      
      vs.rows = vs.rows + 1
      vs.TextMatrix(J, 0) = J
      vs.TextMatrix(J, 1) = RS!NARR
      vs.TextMatrix(J, 2) = ream_tot
      vs.TextMatrix(J, 3) = sheet_tot
      
      RS.MoveNext
    Next
    
    

Else

    k1 = 2
    vs.rows = 2
    If RS.State = 1 Then RS.close
    If cboGodown.text <> "" Then
       RS.Open "select Dates,Narr,OrderNo,ChallanNo,Rec_Reams,Rec_Sheets,Del_Reams,Del_Sheets,Bal_Reams,Bal_Sheets,Printer,aouto from TMPpaperstatement where (Printer='" & cboGodown.text & "' and username='" & UserName & "') order by dates", con
    Else
       RS.Open "select Dates,Narr,OrderNo,ChallanNo,Rec_Reams,Rec_Sheets,Del_Reams,Del_Sheets,Bal_Reams,Bal_Sheets,Printer,aouto from TMPpaperstatement where username='" & UserName & "'  order by aouto", con
    End If
    For J = 1 To RS.RecordCount
      
      vs.rows = vs.rows + 1
      vs.TextMatrix(k1, 0) = RS!dates
      vs.TextMatrix(k1, 1) = RS!NARR
      vs.TextMatrix(k1, 2) = RS!orderNo & ""
      vs.TextMatrix(k1, 3) = RS!ChallanNo & ""
      vs.TextMatrix(k1, 4) = RS!Rec_Reams
      vs.TextMatrix(k1, 5) = RS!Rec_Sheets
      vs.TextMatrix(k1, 6) = RS!Del_Reams
      vs.TextMatrix(k1, 7) = RS!Del_Sheets
      
      
      If k1 <= 2 Then
      
        rm1 = op_ + (RS!Rec_Reams * 500) + RS!Rec_Sheets
        rm1 = rm1 + ((RS!Del_Reams * 500) * -1 + (RS!Del_Sheets * -1))
      
      Else
 
        rm1 = (Val(vs.TextMatrix(k1 - 1, 8)) * 500) + Val(vs.TextMatrix(k1 - 1, 9))
        rm1 = rm1 + (RS!Rec_Reams * 500) + RS!Rec_Sheets
        rm1 = rm1 + ((RS!Del_Reams * 500) * -1 + (RS!Del_Sheets * -1))
      
      End If
 
 
      If Not IsNull(rm1) Then
         Cal_ReamAndSheet 0, rm1
      End If
 
      
      vs.TextMatrix(k1, 8) = ream_tot    'RS!Bal_Reams
      vs.TextMatrix(k1, 9) = sheet_tot             'RS!Bal_Sheets & ""
      
      con.Execute "update TMPpaperstatement set Bal_Reams=" & ream_tot & ",Bal_Sheets=" & sheet_tot & " where userName='" & UserName & "' and aouto=" & RS!aouto & ""
      
      vs.Cell(flexcpFontSize, k1, 1) = 7.5
      k1 = k1 + 1
      RS.MoveNext
    Next
    
End If
'---------------------------------------------------------




cmdprint.Enabled = True

Screen.MousePointer = vbDefault

End Sub
Sub ClosingBal()


Dim rm, ST, vs_raw As Integer

rm = 0
ST = 0


If RS.State = 1 Then RS.close
If cboGodown.text = "" Then
    
     
     RS.Open "select sum(Ream),sum(Sheet) from Order_Qry " & _
    " where " & dateRange & " and " & _
    " Pcode='" & txtPcode & "' and " & stringyear, con, adOpenKeyset, adLockReadOnly


Else
    
     RS.Open "select sum(Ream),sum(Sheet) from Order_Qry " & _
    " where " & dateRange & " and " & _
    " PrinterName='" & cboGodown.text & "' and Pcode='" & txtPcode & "'", con, adOpenKeyset, adLockReadOnly
   
    

End If

If Not IsNull(RS(0)) Then
    rm = rm + (RS(0) * -1)
    ST = ST + (RS(1) * -1)
End If

'------------------------------Data Fatch Paper Rec & Deliver-----------------------------------------

If RS.State = 1 Then RS.close
If cboGodown.text = "" Then
    
    RS.Open "select sum(reams),sum(sheets) from paperstatement where PaperTrans_Deliv='R'" & _
    " and convert(smalldatetime,challan_Date,103)< convert(smalldatetime,'" & FromDate.value & "',103) and pcode='" & txtPcode & "' and " & stringyear & "", con, adOpenKeyset, adLockReadOnly

Else
    
    RS.Open "select sum(reams),sum(sheets) from paperstatement where PaperTrans_Deliv='R' and " & _
    " convert(smalldatetime,challan_Date,103)< convert(smalldatetime,'" & FromDate.value & "',103) and FromGodown='" & cboGodown.text & "' and pcode='" & txtPcode & "' and " & stringyear & "", con, adOpenKeyset, adLockReadOnly

End If
   
If Not IsNull(RS(0)) Then
     rm = rm + RS(0)
     ST = ST + RS(1)
 End If

'------------------------------Paper Delivered---------------------------------------------------------------------------
If RS.State = 1 Then RS.close
If cboGodown.text = "" Then
    RS.Open "select sum(reams),sum(sheets) from paperstatement where PaperTrans_Deliv='D'" & _
    " and convert(smalldatetime,sn_Date,103)< convert(smalldatetime,'" & FromDate.value & "',103) and pcode='" & txtPcode & "' and " & stringyear & "", con, adOpenKeyset, adLockReadOnly

Else
    RS.Open "select sum(reams),sum(sheets) from paperstatement where PaperTrans_Deliv='D' and " & _
    "  convert(smalldatetime,sn_Date,103)< convert(smalldatetime,'" & FromDate.value & "',103)and FromGodown='" & cboGodown.text & "' and pcode='" & txtPcode & "' and " & stringyear & "", con, adOpenKeyset, adLockReadOnly
End If

If Not IsNull(RS(0)) Then
     rm = rm + (RS(0) * -1)
     ST = ST + (RS(1) * -1)
 End If



Cal_ReamAndSheet rm, ST

txtReam_op.text = ream_tot
txtSheet_op.text = sheet_tot



End Sub

Private Sub Form_Load()

 Me.Left = 50
 Me.top = 50
 Me.Width = 11450
 Me.Height = 7900

'If session = "2019-20" Then
'FromDate.value = Format("04/06/2019", "dd/MM/yyyy")
'Else
FromDate.value = Format(from_date, "dd/MM/yyyy")
'End If

toDate.value = Format(to_date, "dd/MM/yyyy")


If RS.State = 1 Then RS.close
RS.Open "select Godwn as [Binder Name],Address,id from Godownmaster where " & stringyear & " and len(Godwn)>3 order by Godwn", con, adOpenStatic, adLockReadOnly

While RS.EOF = False
  cboGodown.AddItem RS(0)
  cboGodown.ItemData(cboGodown.NewIndex) = RS!id
  RS.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "Select size1 from sizeMaster where " & stringyear & " and size1 <> '' order by size1", con
While RS.EOF = False
   cboPaperSize.AddItem RS(0)
   RS.MoveNext
Wend

vs_ini

BackColorFrom Me, 1

If LCase(UserName) = "admin" Then
   cmdOp.Visible = True
Else
   cmdOp.Visible = False
End If

End Sub

