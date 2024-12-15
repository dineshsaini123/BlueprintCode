VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMonthlyDeb 
   Caption         =   "Monthly Book Debts"
   ClientHeight    =   9084
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14208
   Icon            =   "frmMonthlyDeb.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9084
   ScaleWidth      =   14208
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export to Excel"
      Height          =   555
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   60
      Width           =   1260
   End
   Begin VB.CommandButton cmdref 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   555
      Left            =   2925
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   60
      Width           =   1260
   End
   Begin VB.CheckBox Check1_print 
      Caption         =   "Print Exceed Column"
      Height          =   240
      Left            =   5400
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   8925
      TabIndex        =   20
      Top             =   8415
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   6825
      TabIndex        =   19
      Top             =   8415
      Width           =   2085
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   4080
      TabIndex        =   18
      Top             =   8430
      Width           =   2715
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1425
      TabIndex        =   17
      Top             =   8445
      Width           =   2625
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   150
      TabIndex        =   16
      Top             =   8445
      Width           =   1275
   End
   Begin VB.TextBox txtdays 
      Height          =   285
      Left            =   4065
      TabIndex        =   14
      Text            =   "120"
      Top             =   705
      Width           =   720
   End
   Begin Crystal.CrystalReport cr 
      Left            =   13050
      Top             =   450
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   555
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   60
      Width           =   1260
   End
   Begin MSComCtl2.DTPicker dates 
      Height          =   330
      Left            =   1155
      TabIndex        =   11
      Top             =   105
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   572
      _Version        =   393216
      Format          =   73465857
      CurrentDate     =   39127
   End
   Begin VB.TextBox txtamt2 
      Height          =   270
      Left            =   12600
      TabIndex        =   10
      Text            =   "0"
      Top             =   75
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtAmt1 
      Height          =   270
      Left            =   12825
      TabIndex        =   9
      Text            =   "0"
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   555
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   1260
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   555
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1260
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   288
      Left            =   10530
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   345
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7260
      Left            =   150
      TabIndex        =   0
      Top             =   1065
      Width           =   13830
      _cx             =   24395
      _cy             =   12806
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12582847
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483647
      BackColorBkg    =   16777215
      BackColorAlternate=   12582847
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
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
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Editable        =   0
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
   Begin MSMask.MaskEdBox date1 
      Height          =   285
      Left            =   10560
      TabIndex        =   4
      Top             =   75
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   508
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   285
      Left            =   11550
      TabIndex        =   5
      Top             =   75
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1863
      _ExtentY        =   508
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTotal 
      Height          =   285
      Left            =   11790
      TabIndex        =   25
      Top             =   8415
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label lblno 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   9675
      TabIndex        =   21
      Top             =   150
      Width           =   600
   End
   Begin VB.Label Label3 
      Caption         =   "With in days"
      Height          =   285
      Left            =   2760
      TabIndex        =   15
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Current Date"
      Height          =   285
      Left            =   165
      TabIndex        =   12
      Top             =   135
      Width           =   1110
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Amt Greater Than"
      Height          =   195
      Index           =   1
      Left            =   8310
      TabIndex        =   8
      Top             =   1635
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   2
      Left            =   8670
      TabIndex        =   7
      Top             =   1575
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Amt. Less Than"
      Height          =   315
      Index           =   3
      Left            =   8310
      TabIndex        =   6
      Top             =   1395
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "frmMonthlyDeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Dim current_Last As String
Dim RS As Recordset
Function rsets(ST As String, length As Integer) As String
   
    Dim kk As String
            kk = Trim(ST)
            If Len(kk) < length Then
                Do While Not Len(kk) = length
                    kk = " " + kk
                Loop
            End If
            If Len(kk) > length Then
                Do While Not Len(kk) = length
                    kk = Mid$(kk, 0, Len(kk) - 1)
                Loop
            End If
        rsets = kk
End Function


Private Sub alpha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sendkeys "{TAB}"
End If
End Sub

Private Sub cmdPrint_Click()
   
   
   DSNNew
   
   Screen.MousePointer = vbHourglass
   
   con.Execute "delete from winrptTmp where uid=" & UId & ""
   con.Execute "update MonthlyDetHead set  [1]='" & IIf(Text1.text = "", "-", Text1.text) & "',[2]='" & IIf(Text2.text = "", "-", Text2.text) & "',[3]='" & IIf(Text3.text = "", "-", Text3.text) & "' ,[4]='" & IIf(Text4.text = "", "-", Text4.text) & "', [5]= '" & IIf(Text5.text = "", "-", Text5.text) & "'"
   
   For I = 1 To vs.rows - 1
      If vs.TextMatrix(I, 1) <> "" Then
      If Val(vs.TextMatrix(I, 2)) = 0 And Val(vs.TextMatrix(I, 4)) = 0 And Val(vs.TextMatrix(I, 6)) = 0 Then
      Else
      con.Execute "INSERT INTO winrptTmp(Description,op,Party,Receipt,dr,Payment,cr,Balance,uid) values('" & vs.TextMatrix(I, 1) & "'," & Val(vs.TextMatrix(I, 2)) & ",'" & vs.TextMatrix(I, 3) & "'," & Val(vs.TextMatrix(I, 4)) & ",'" & Val(vs.TextMatrix(I, 5)) & "'," & Val(vs.TextMatrix(I, 6)) & ",'" & vs.TextMatrix(I, 7) & "'," & I & "," & UId & ")"
      End If
      End If
   Next
   
  Screen.MousePointer = vbDefault
  
  Commandshow.Enabled = True
  
   
   MsgBox "Want to View ?", vbInformation
   
   
    CR.Reset
    'If Check1_print.value = 1 Then
    'cr.ReportFileName = st1 & "\" & main.directory & "\Monthlydebt.rpt"
    'Else
    CR.ReportFileName = rptPath & "/MonthlydebtNew.rpt"
    'End If
    
    If RS.State = 1 Then RS.close
    RS.Open "select * from MonthlyDetHead", con
    If RS.EOF = False Then
     CR.Formulas(0) = "with90='" & RS(3) & "'"
     CR.Formulas(1) = "exeed='" & RS(4) & "'"
     CR.Formulas(4) = "srno='" & RS(0) & "'"
     CR.Formulas(5) = "des='" & RS(1) & "'"
     CR.Formulas(6) = "total='" & RS(2) & "'"
    End If
    CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    CR.ReplaceSelectionFormula "{Winrpt.uid}=" & UId & ""
    CR.Formulas(3) = "dates='" & dates.value & "'"
    CR.WindowShowRefreshBtn = True
    CR.WindowShowPrintBtn = True
    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowSearchBtn = True
    CR.WindowState = crptMaximized
    CR.WindowShowExportBtn = True
    CR.Action = 1
    
    
    

   
End Sub

Private Sub cmdref_Click()
  con.Execute "Delete  from subledgertrail"
  con.Execute ("delete  from treport_tmp")
  
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
Dim Q1, q2, J As Double

Dim b1 As Boolean

b1 = False


c = 1
r = 1



row_ = 1
col_ = 1
   
    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    xlSheet.Cells(1, 1).value = "Bilty Return Status "
    
    For I = 0 To vs.rows - 1
        For J = 0 To vs.Cols - 1
               xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
              col_ = col_ + 1
        Next
        row_ = row_ + 1
        col_ = 1
    Next
    

MsgBox "final...."


Exit Sub
Screen.MousePointer = vbDefault
err:
MsgBox err.DESCRIPTION



End Sub

Private Sub COMBOGENLEDGER_Change()
    If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
    'Combosubledger.Clear
    If Not RS.BOF Then
        Do While Not RS.EOF
            Combosubledger.AddItem Trim(RS!subledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    
End Sub


Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   sendkeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
   sendkeys "{DOWN}"
   sendkeys "{tab}"
End If

End Sub

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.text) <> "" Then
        Set RS = New ADODB.Recordset
        RS.Open "select * from gledger where slf=true", con, adOpenStatic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.Find "gledger='" + Trim(COMBOGENLEDGER.text) + "'"
            If RS.EOF Then
                COMBOGENLEDGER.SetFocus
            End If
        Else
            COMBOGENLEDGER.SetFocus
        End If
        RS.close
    End If
End Sub

Private Sub Combosubledger_GotFocus()
    If Trim(COMBOGENLEDGER.text) = "" Then
        COMBOGENLEDGER.SetFocus
    End If
End Sub

Private Sub Combosubledger_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   sendkeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
   sendkeys "{Down}"
   sendkeys "{tab}"
End If
End Sub

Private Sub Combosubledger_LostFocus()
If Trim(Combosubledger.text) <> "" Then
    If Trim(COMBOGENLEDGER.text) <> "" Then
        If RS.State = 1 Then
            RS.close
        End If
        RS.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.close
    Else
        Combosubledger.text = ""
    End If
End If
End Sub
Sub ALPHAB()

If RS.State = 1 Then RS.close
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
con.Execute ("delete  from treport_tmp")



Dim rs1 As New ADODB.Recordset
Dim Balance As Double
Dim OPBALANCE As Double
Dim SDamount As Double
Dim SCamount As Double
Dim RsT As New ADODB.Recordset
Dim viewsubledger As Boolean
Dim date1_Last As String
Dim date2_Last As String

viewsubledger = False
Balance = 0
OPBALANCE = 0

'OPENINGSUBLEDGERS



DoEvents

If session = "2024-25" Then
   
   If Val(txtdays) > 0 Then
      date1_Last = DateAdd("d", Val(txtdays) * -1, dates.value)
   Else
      date1_Last = "01/12/2023"
   End If
   
   
   If (DateValue(date1_Last) < DateValue(from_date)) Then
      date2_Last = "31/03/2024"
      
      con.Execute "exec balance_120 '" & date1_Last & "','" & date2_Last & "','" & UId & "'"
      con_LAST.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "last"
   Else
      con.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "current"
   End If
   
   
End If


If session = "2023-24" Then
   
   If Val(txtdays) > 0 Then
      date1_Last = DateAdd("d", Val(txtdays) * -1, dates.value)
   Else
      date1_Last = "01/12/2022"
   End If
   
   
   If (DateValue(date1_Last) < DateValue(from_date)) Then
      date2_Last = "31/03/2023"
      
      con.Execute "exec balance_120 '" & date1_Last & "','" & date2_Last & "','" & UId & "'"
      con_LAST.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "last"
   Else
      con.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "current"
   End If
   
End If

If session = "2022-23" Then
   
   If Val(txtdays) > 0 Then
      date1_Last = DateAdd("d", Val(txtdays) * -1, dates.value)
   Else
      date1_Last = "01/12/2021"
   End If
   
   
   If (DateValue(date1_Last) < DateValue(from_date)) Then
      date2_Last = "31/03/2022"
      
      con.Execute "exec balance_120 '" & date1_Last & "','" & date2_Last & "','" & UId & "'"
      con_LAST.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "last"
   Else
      con.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "current"
   End If
   
End If


If session = "2020-21" Then
   
   If Val(txtdays) > 0 Then
      date1_Last = DateAdd("d", Val(txtdays) * -1, dates.value)
   Else
      date1_Last = "01/12/2019"
   End If
   
   
   If (DateValue(date1_Last) < DateValue(from_date)) Then
      date2_Last = "31/03/2020"
      con.Execute "exec balance_120 '" & date1_Last & "','" & date2_Last & "','" & UId & "'"
      con_LAST.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "last"
   Else
      con.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "current"
   End If
   
End If

If session = "2021-22" Then
   
   If Val(txtdays) > 0 Then
      date1_Last = DateAdd("d", Val(txtdays) * -1, dates.value)
   Else
      date1_Last = "01/12/2020"
   End If
   
   
   If (DateValue(date1_Last) < DateValue(from_date)) Then
      date2_Last = "31/03/2021"
      
      con.Execute "exec balance_120 '" & date1_Last & "','" & date2_Last & "','" & UId & "'"
      con_LAST.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "last"
   Else
      con.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "current"
   End If
   
End If


If session = "2019-20" Then
   
   If Val(txtdays) > 0 Then
   date1_Last = DateAdd("d", Val(txtdays) * -1, dates.value)
   Else
   date1_Last = "01/12/2018"
   End If
   
   
   If (DateValue(date1_Last) < DateValue(from_date)) Then
      date2_Last = "31/03/2019"
      con.Execute "exec balance_120 '" & date1_Last & "','" & date2_Last & "','" & UId & "'"
      con_LAST.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "last"
   Else
      con.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "current"
   End If
   
   
   
End If


If session = "2018-19" Then
   
   If Val(txtdays) > 0 Then
   date1_Last = DateAdd("d", Val(txtdays) * -1, dates.value)
   Else
   date1_Last = "01/12/2017"
   End If
   
   
   
   
   
   If (DateValue(date1_Last) < DateValue(from_date)) Then
      
      con.Execute "exec balance_120 '" & date1_Last & "','" & date2_Last & "','" & UId & "'"
      con_LAST.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "last"
   Else
      con.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "current"
   End If
   
End If



If session = "2017-2018" Then
   
   If Val(txtdays) > 0 Then
   date1_Last = DateAdd("d", Val(txtdays) * -1, dates.value)
   Else
   date1_Last = "01/12/2016"
   End If
   date2_Last = "31/03/2017"
   
   
   
   If (DateValue(date1_Last) < DateValue(from_date)) Then
      con.Execute "exec balance_120 '" & date1_Last & "','" & date2_Last & "','" & UId & "'"
      con_LAST.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "last"
   Else
      con.Execute "exec OPbalance_120 '" & date1_Last & "'"
      current_Last = "current"
   End If
   
   
End If



con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where   genledger ='" + Trim(COMBOGENLEDGER.text) + "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where   genledger ='" + Trim(COMBOGENLEDGER.text) + "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid) SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3  , " & UId & " FROM INVOICEA  where   genledger ='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103) "
con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid) SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & " FROM CASHA   where   genledger='" + Trim(COMBOGENLEDGER.text) + "'    and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)  "
con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "  FROM CASHA  where  genledger='" + Trim(COMBOGENLEDGER.text) + "'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)  AND CASHA.BAA <>0  "
con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & "  FROM CREDITA  where   genledger='" + Trim(COMBOGENLEDGER.text) + "'    and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2) + "',103) "
con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, ''  , " & UId & " From DNFA  where   Pgld ='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2) + "',103) ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & " From CNF1A where Pgld='" + Trim(COMBOGENLEDGER.text) + "'  and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '', " & UId & "  From DNFB  where gld='" + Trim(COMBOGENLEDGER.text) + "'   and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2) + "',103)   ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
con.Execute "INSERT INTO treport_tmp ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & " From CNF1B where gld='" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"


''con.Execute "delete from treport_tmp"

If current_Last = "last" Then
  If rs1.State = 1 Then rs1.close
  rs1.Open "SELECT '" + Trim(COMBOGENLEDGER.text) + "' as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId from subledgertrail_net GROUP BY SUBLEDGER", con_LAST
  While rs1.EOF = False
     con.Execute "insert into treport_tmp (Genledger,Subledger,openingbalance,userid) values('" & rs1!genled & "','" & rs1!subledger & "'," & rs1!opcr & ", '" & UId & "')"
     rs1.MoveNext
  Wend
  ''con.Execute "insert into treport_tmp (Genledger,Subledger,openingbalance,userid) SELECT '" + Trim(COMBOGENLEDGER.Text) + "'as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId from subledgertrail_net GROUP BY SUBLEDGER"
Else
  If rs1.State = 1 Then rs1.close
  rs1.Open "SELECT '" + Trim(COMBOGENLEDGER.text) + "' as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId from subledgertrail_net GROUP BY SUBLEDGER", con
  While rs1.EOF = False
     con.Execute "insert into treport_tmp (Genledger,Subledger,openingbalance,userid) values('" & rs1!genled & "','" & rs1!subledger & "'," & rs1!opcr & ", '" & UId & "')"
     rs1.MoveNext
  Wend
End If

con.Execute "update treport_tmp set header='" & dates.value & "'"
main.reportname = "Sub. Ledger"

End Sub
Sub showdatainGrid()

    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs_op As ADODB.Recordset
    
    Dim dr, CR, dr1, cr1, op As Double
    Dim dr_op, cr_op As Double
    Dim bal, Total As Double
    
    Total = 0
    dr = 0
    CR = 0
    op = 0
    dr_op = 0
    cr_op = 0
    bal = 0
    
    kk = 1

    Set rs_op = New ADODB.Recordset
    date1_Last = DateAdd("d", Val(txtdays) * -1, dates.value)
   
    con.Execute "update treport_tmp set tab= DATEDIFF(day,vdate,convert(datetime, header,103)+1)  where vdate  is not null "
    con.Execute "update treport_tmp set tab=" & 360 & " where tab=" & 0 & ""

    Dim K As Integer
    K = 0
    vs.Cols = 9
    vs.rows = 2
    If RS.State = 1 Then RS.close
    RS.Open "select distinct(subledger) from treport_tmp where genledger='SUNDRY DEBTORS' order by subledger", con
    If RS.EOF = False Then
       DoEvents
       lblno.Caption = RS.RecordCount
    End If
    
    For k2 = 1 To RS.RecordCount
        
        If rs1.State = 1 Then rs1.close
        'rs1.Open "select sum(OpeningBalance) from treport_tmp where (genledger='SUNDRY DEBTORS' and subledger='" & RS(0) & "')", con
        If current_Last = "last" Then
           rs1.Open "SELECT sum(YEAROPENING),SUM (OPAMOUNTDEBIT),SUM(OPAMOUNTCREDIT) from subledgertrail_net where subledger='" & RS(0) & "'", con_LAST
        Else
           rs1.Open "SELECT sum(YEAROPENING),SUM (OPAMOUNTDEBIT),SUM(OPAMOUNTCREDIT) from subledgertrail_net where subledger='" & RS(0) & "'", con
        End If
        If rs1.RecordCount > 0 Then
           op = ((IIf(IsNull(rs1(0)), 0, rs1(0)) + (IIf(IsNull(rs1(1)), 0, rs1(1)) - IIf(IsNull(rs1(2)), 0, rs1(2)))))
           bal = IIf(IsNull(rs1(0)), 0, rs1(0))
        Else
        op = 0
        End If
        
        

        
        
         If RS(0) = "K2108 KUMAR BOOK AGENCY (EM), MEERUT" Then
            
          '  MsgBox "s"
         End If
        
       
       
        If rs1.State = 1 Then rs1.close
        rs1.Open "select sum(ad) from treport_tmp where (genledger='SUNDRY DEBTORS' and subledger='" & RS(0) & "') and dorc='D' and vdate >= convert(smalldatetime,'" + Trim(date1_Last) + "',103)   and vdate <=convert(smalldatetime,'" + Trim(date2.text) + "',103) ", con
        If Not IsNull(rs1(0)) Then
        dr = rs1(0)
        Else
        dr = 0
        End If

        If rs1.State = 1 Then rs1.close
        rs1.Open "select sum(ad) from treport_tmp where (genledger='SUNDRY DEBTORS' and subledger='" & RS(0) & "') and dorc='C' and vdate >= convert(smalldatetime,'" + Trim(date1_Last) + "',103)   and vdate <=convert(smalldatetime,'" + Trim(date2.text) + "',103) ", con
        If Not IsNull(rs1(0)) Then
        CR = rs1(0)
        Else
        CR = 0
        End If

        If RS(0) = "K2108 KUMAR BOOK AGENCY (EM), MEERUT" Then
        '   MsgBox "a"
        End If

        
       DoEvents
       Total = Total + (dr - CR) + op
       DoEvents

       If IIf((Val(((dr - CR) + op)) > 0), "Dr", "Cr") = "Dr" Then
        
        
        K = K + 1
        vs.rows = vs.rows + 1
        vs.TextMatrix(K, 0) = K
        vs.TextMatrix(K, 1) = RS(0)
        '''vs.TextMatrix(K, 2) = (op + (dr_op - cr_op))
        vs.TextMatrix(K, 2) = Round(op, 2)
        vs.TextMatrix(K, 6) = Round((dr - CR) + op, 2)
        
        
        dr = 0: CR = 0: cr1 = 0: dr1 = 0
        
       
    
    
    If rs2.State = 1 Then rs2.close
    rs2.Open "select * from treport_tmp where genledger='SUNDRY DEBTORS' and subledger='" & RS(0) & "'", con
    If rs2.EOF = False Then
    While rs2.EOF = False
        
        If rs2.Fields("tab").value <= Val(txtdays) Then
            If rs2.Fields("dorc").value = "D" Then
                dr = dr + rs2!ad
            
            End If
            
            If rs2.Fields("dorc").value = "C" Then
                dr1 = dr1 + rs2!ad
            End If
        
        
        End If
        '''Else
        
        
        
         rs2.MoveNext
    Wend
    End If
        

    
    vs.TextMatrix(K, 4) = Round(dr, 2)
    vs.TextMatrix(K, 7) = Round(dr, 2)
    vs.TextMatrix(K, 5) = dr1
    vs.TextMatrix(K, 8) = vs.TextMatrix(K, 6) - vs.TextMatrix(K, 4)
    
    
    If Val(vs.TextMatrix(K, 8)) < 0 Then
       vs.TextMatrix(K, 7) = Val(vs.TextMatrix(K, 7)) + Val(vs.TextMatrix(K, 8))
       vs.TextMatrix(K, 8) = 0
    End If


 
End If

DoEvents
DoEvents

lblno.Caption = (lblno.Caption) - 1
lblTotal.Caption = Total


kk = kk + 1
        
'        If kk >= 200 Then
'           GoTo aaaaa:
'        End If

RS.MoveNext
Next

'aaaaa:

    'vs.FormatString = "SrNo|Name of Debitor|Balance|Dr/Cr|With in " & txtdays.Text & " Days|Dr/Cr||"
    vs.FormatString = "SrNo|Name of Debitor|OP.Bal|Dr/Cr|Sale With in " & txtdays.text & " Days|Rec. With in " & txtdays.text & " Days|Outstanding|Less Then " & txtdays.text & "|More Then " & txtdays.text & ""
    
    vs.ColWidth(0) = 600
    vs.ColWidth(1) = 3800
    vs.ColWidth(2) = 1400
    vs.ColWidth(3) = 0
    vs.ColWidth(4) = 1700
    vs.ColWidth(5) = 1700
    vs.ColWidth(6) = 1200
    vs.ColWidth(7) = 1200
    vs.ColWidth(8) = 1200


'''
End Sub
Private Sub CommandReturn_Click()
    'MainMenu.Toolbar1.Visible = false
    Unload Me
End Sub
Private Sub Commandshow_Click()
 Commandshow.Enabled = False
 DoEvents
 
 'date1.Text = DateAdd("D", dates, -120)
 
 con.Execute "Delete  from subledgertrail"
 con.Execute ("delete  from treport_tmp")
 
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 
 Screen.MousePointer = vbHourglass
 
 ALPHAB
 Commandshow.Enabled = False
 cmdprint.Enabled = True
 
 showdatainGrid
 Commandshow.Enabled = True
 
 MsgBox "updated....", vbInformation
 
 Screen.MousePointer = vbDefault
 
End Sub
Function Genrpt()

    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim dr, CR, dr1, cr1
    Dim op As Double
    Dim closDr As Double
    Dim closCr As Double
    op = 0
    closDr = 0
    closCr = 0
    dr = 0
    CR = 0
    dr1 = 0
    cr1 = 0
    Screen.MousePointer = vbHourglass
    '--------------------------------------------------------------
    Dim Balance As Double
    Balance = 0
  ' opening balance start
    '--------------------------------------------------------------
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim xtemp As String
    Dim header As String
    Dim DateHeader As String
    DateHeader = ""
    header = ""
    Dim trs As ADODB.Recordset
    Dim Pno As Integer
    Set trs = New ADODB.Recordset
    con.Execute "delete from Winrpt where uid=" & UId & ""
    paperWidth = 150
        T1 = 20
        T2 = 30
        T3 = 40
        T4 = 50
        T5 = 80
        T6 = 90
        T7 = 110
        T8 = 115
        header = COMBOGENLEDGER.text
        
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim b As Boolean
        Dim Cr10 As Double
        Dim Dr10 As Double
        
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Dim FooterYes As Integer
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        Dim drhead As String
        FooterYes = False
        main.reportdata
        main.repors.Find "reportname='" + Trim(main.reportname) + "'"
        MaxLine = main.repors!totalline
        If main.repors!comp = True Then
            paperWidth = Int(main.repors!totalcolumn * 1.75)
        Else
            paperWidth = main.repors!totalcolumn
        End If
        paperWidth = 136
        MaxLine = 72
        Open "" + VB.App.Path + "\vipin.txt" For Output As #1
        Line = 0
        Pno = 1
header:
        Dim I As Integer
        For I = 1 To main.repors!TopMargin
            Print #1, ""
            Line = Line + 1
        Next
       If FooterYes = True Then
             Print #1, ""
           ' Print #1, Tab(LEFTM); repli("-", paperWidth)
            Line = Line + 1
            Do While Line < 72
                    Print #1, " "
                    Line = Line + 1
            Loop
            Line = 0
            FooterYes = False
       End If
       Dim bb As Boolean
       If kkk.State = 1 Then kkk.close
       CNSetup
       kkk.Open "select * from setup1", con, adOpenDynamic, adLockReadOnly, adCmdText
       If Not kkk.BOF Then
            Print #1, ""
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(119); "Session: " & ses
            Print #1, Tab(125); "Page No:  " & Pno
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname)
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            Line = Line + 6
        End If
        If trs.State = 1 Then
            trs.close
        End If
        xstr = date1.text & " To " & date2.text
        Print #1, Chr(27) + Chr(14); Tab((69 - Len(Trim(header))) / 2); Trim(header); Chr(27) + Chr(15)
        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        '=============== Head Name
        Print #1, Tab(3); "PARTY NAME"; Tab(T2 + 20); "OPENING"; Tab(T3 + 25 + LEFTM); "YTD DEBIT"; Tab(T5 - 2 + LEFTM); "YTC CREDIT"; Tab(T6 + 4 + LEFTM); "NET DEBIT"; Tab(T7 + 3 + LEFTM); "NET CREDIT"
        '---------------
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Line = Line + 5
        'trs.Close
        If called1 = True Then GoTo printagain1
        If rs1.State = 1 Then rs1.close
        rs1.Open "select subledger,yearopening from subledgertrail order by subledger", con
        If Not rs1.BOF Then
        pg.Max = rs1.RecordCount
        Do While Not rs1.EOF
        If pg.value >= pg.Max Then
           pg.value = 0
        Else
           pg.value = pg.value + 1
        End If
        dr = 0
        CR = 0
        Cr10 = 0
        Dr10 = 0
        If RS.State = 1 Then RS.close
        RS.Open "select sum(ad) from treport_tmp where subledger='" + Trim(rs1!subledger) + "' and dorc='D'", con
        If RS(0) <> 0 Then
        dr = RS(0)
        End If
        If RS.State = 1 Then RS.close
        RS.Open "select sum(ad) from treport_tmp where subledger='" + Trim(rs1!subledger) + "' and dorc='C'", con
        If RS(0) <> 0 Then
        CR = RS(0)
        End If
        'dr1 = dr1 + dr
        'cr1 = cr1 + cr
        CR = (-1 * CR)
        op = op + rs1!YEAROPENING
        If rs1.Fields("yearopening").value < 0 Then
           Cr10 = CR + rs1.Fields("yearopening").value
         Else
           Dr10 = dr + rs1.Fields("yearopening").value
        End If
        Dr10 = Dr10 + Cr10
        If Dr10 < 0 Then
           Dr10 = (-1 * Dr10)
        End If
        
    If Val(Dr10) >= Val(txtAmt1.text) And Val(Dr10) < Val(txtamt2.text) Then
        '==================================================================
abc:
        Print #1, Tab(3); Trim(rs1!subledger);
        If Val(rs1!YEAROPENING) < 0 Then
           Print #1, Tab(T1 + 25); Space(12 - Len(Format(Str((-1 * rs1!YEAROPENING)), "0.00"))) & Format(Str((-1 * rs1!YEAROPENING)), "0.00") & " Cr";
           drhead = "Cr"
        Else
           drhead = "Dr"
           Print #1, Tab(T1 + 25); Space(12 - Len(Format(Str(rs1!YEAROPENING), "0.00"))) & Format(Str(rs1!YEAROPENING), "0.00") & " Dr";
        End If
        Print #1, Tab(T2 + 32); Space(12 - Len(Format(Str(dr), "0.00"))) & Format(Str(dr), "0.00");
        Print #1, Tab(T3 + 36); Space(12 - Len(Format(Str((-1 * CR)), "0.00"))) & Format(Str((-1 * CR)), "0.00");
        cr1 = cr1 + (-1 * CR)
        dr1 = dr1 + dr
        If rs1.Fields("yearopening").value < 0 Then
           CR = CR + rs1.Fields("yearopening").value
         Else
           dr = dr + rs1.Fields("yearopening").value
        End If
        closDr = closDr + dr
        closCr = closCr + CR
        Print #1, Tab(T4 + 41); Space(12 - Len(Format(Str((dr)), "0.00"))) & Format(Str((dr)), "0.00");
        Print #1, Tab(T5 + 31); Space(12 - Len(Format(Str(((-1 * CR))), "0.00"))) & Format(Str(((-1 * CR))), "0.00");
        '-------------------------------------------------------------
         con.Execute "insert into winrpt(Party,op,Receipt,Payment,closing,closing1,dr,Description,FromDate,toDate,uid) values('" & Trim(rs1!subledger) & "'," & rs1!YEAROPENING & "," & dr1 & "," & cr1 & "," & dr & "," & (-1 * CR) & ",'" & drhead & "','" & COMBOGENLEDGER.text & "','" & date1.text & "','" & date2.text & "'," & UId & ")"
        '=============================================================
         Line = Line + 1
         If Line > MaxLine - 8 Then
         called1 = True
printnext:
         Pno = Pno + 1
         FooterYes = True
         GoTo header
printagain1:
         called1 = False
         End If
      ElseIf Val(txtAmt1.text) = 0 Or Val(txtamt2.text) = 0 Then
         GoTo abc:
      End If
         If Not rs1.EOF Then
         rs1.MoveNext
         End If
         Loop
printfooter:
              Print #1, Tab(LEFTM); repli("-", paperWidth)
              ''''------------------
              If op < 0 Then
                Print #1, Tab(T1 + 25); Space(12 - Len(Format(Str((-1 * op)), "0.00"))) & Format(Str((-1 * op)), "0.00") & " Cr";
              Else
                Print #1, Tab(T1 + 25); Space(12 - Len(Format(Str(op), "0.00"))) & Format(Str(op), "0.00") & " Dr";
              End If
              Print #1, Tab(T2 + 32 + LEFTM); rsets(Trim(Format(Str(dr1), "0.00")), 12);
              Print #1, Tab(T3 + 36 + LEFTM); rsets(Trim(Format(Str(cr1), "0.00")), 12);
              Print #1, Tab(T4 + 39); Space(14 - Len(Format(Str((closDr)), "0.00"))) & Format(Str((closDr)), "0.00");
              Print #1, Tab(T5 + 29); Space(14 - Len(Format(Str((closCr)), "0.00"))) & Format(Str(Abs(closCr)), "0.00");
              ''--------------------
              Line = Line + 1
              Do While Line <= 72
                     Print #1, " "
                     Line = Line + 1
              Loop
        End If
        Close #1
        Screen.MousePointer = vbDefault
End Function

Private Sub date1_KeyPress(KeyAscii As Integer)
  
If KeyAscii = 13 Then
    date2.SetFocus
End If
End Sub

Private Sub date1_LostFocus()
    If Trim(date1.text) <> "" Then
        If Not checkdate(Trim(date1.text), date1) Then
            date1.SetFocus
            End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sendkeys "{TAB}"
End If

End Sub

Private Sub date2_LostFocus()
    If Trim(date2.text) <> "" Then
        If Not checkdate(Trim(date2.text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub


Private Sub Dates_LostFocus()
date2.text = dates.value
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 '   SendKeys "{TAB}"
End If

End Sub
Private Sub Form_Load()


Set RS = New ADODB.Recordset
RS.Open "select * from MonthlyDetHead", con
If RS.EOF = False Then
 Text1.text = RS(0)
 Text2.text = RS(1)
 Text3.text = RS(2)
 Text4.text = RS(3)
 Text5.text = RS(4)
End If

con.Execute "delete  from treport_tmp"
con.Execute "Delete  from subledgertrail"

Me.top = 25
Me.Left = 20

Set RS = New ADODB.Recordset
RS.Open "select gledger from GLEDGER where gledger  = '" & "SUNDRY DEBTORS" & "'", con
While RS.EOF = False
  COMBOGENLEDGER.AddItem RS.Fields(0).value
  RS.MoveNext
Wend
    
Set RS = New ADODB.Recordset
RS.Open "setup1", con, adOpenStatic, adLockReadOnly, adCmdTable
CNSetup
date1.text = RS!yarfrom
date2.text = RS!yarto
RS.close

COMBOGENLEDGER.ListIndex = 0

dates.value = Date

date2.text = dates.value

End Sub
Sub xx()


End Sub
Sub OPENINGSUBLEDGERS()

         
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , YEAROPENING,  0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "'", p, adCmdText

        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER))  " _
        & " where invoiceA.setupid=" & setupid & " and invoiceA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "'and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
        & " where  cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and  SLEDGER.SUBLEDGER <>'CASH PARTY' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
        & " where cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
        & " where  credita.setupid=" & setupid & " and credita.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE vouchers.setupid=" & setupid & " and vouchers.fyear='" & session & "' and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE vouchers.setupid=" & setupid & " and vouchers.fyear='" & session & "' and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE cnf1a.setupid=" & setupid & " and cnf1a.fyear='" & session & "' and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        'and ReflectInAcc=0
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE cnf1a.setupid=" & setupid & " and cnf1a.fyear='" & session & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        ' and ReflectInAcc=0
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE dnfa.setupid=" & setupid & " and dnfa.fyear='" & session & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE dnfa.setupid=" & setupid & " and dnfa.fyear='" & session & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE  CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "' and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE dnfB.setupid=" & setupid & " and dnfB.fyear='" & session & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE dnfB.setupid=" & setupid & " and dnfB.fyear='" & session & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
  
  
  
End Sub
Private Sub txtdays_LostFocus()
    Text4.text = "With " & txtdays.text & " Days"
    Text5.text = "Exceed " & txtdays.text & " Days"
End Sub
