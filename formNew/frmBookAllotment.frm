VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBookAllotment 
   Caption         =   "Rep. Wise Specimen Allotment"
   ClientHeight    =   9144
   ClientLeft      =   60
   ClientTop       =   516
   ClientWidth     =   11592
   Icon            =   "frmBookAllotment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9144
   ScaleWidth      =   11592
   Begin VB.CheckBox Check1_UpdateFromExcel 
      Caption         =   "Update Specimen Allotment from Excel File"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5940
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export/Import Excel File"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   5940
      TabIndex        =   11
      Top             =   585
      Visible         =   0   'False
      Width           =   5556
      Begin VB.CommandButton cmdGenerateFile 
         Caption         =   "&Generate Excel  File"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   900
         TabIndex        =   16
         Top             =   855
         Width           =   915
      End
      Begin VB.CommandButton cmdUpdateSp 
         Caption         =   "&Update Specime"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   90
         TabIndex        =   15
         Top             =   855
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5085
         Top             =   900
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "&Import File"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4776
         TabIndex        =   14
         Top             =   405
         Width           =   735
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Left            =   90
         TabIndex        =   13
         Top             =   495
         Width           =   4692
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00B8E4F1&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   1470
      ScaleHeight     =   756
      ScaleWidth      =   3252
      TabIndex        =   7
      Top             =   1005
      Width           =   3255
      Begin VB.CommandButton cmdAddNewRep 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add New Rep.App."
         Height          =   645
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton CommandPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Excel"
         Height          =   645
         Left            =   1035
         Picture         =   "frmBookAllotment.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   45
         Width           =   990
      End
      Begin VB.CommandButton Commandhelp 
         Caption         =   "Help"
         Height          =   495
         Left            =   -720
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton Commandsave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sa&ve"
         Height          =   645
         Left            =   60
         Picture         =   "frmBookAllotment.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton CommandReturn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Return"
         Height          =   645
         Left            =   2070
         Picture         =   "frmBookAllotment.frx":17D4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   1005
      End
   End
   Begin VB.ComboBox cborep 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      ItemData        =   "frmBookAllotment.frx":23B8
      Left            =   8388
      List            =   "frmBookAllotment.frx":23BA
      TabIndex        =   0
      Top             =   252
      Visible         =   0   'False
      Width           =   540
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6420
      Left            =   96
      TabIndex        =   3
      Top             =   2076
      Width           =   11292
      _cx             =   19918
      _cy             =   11324
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   12648447
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   -2147483647
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBookAllotment.frx":23BC
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
   End
   Begin MSComCtl2.DTPicker fdate 
      Height          =   330
      Left            =   1500
      TabIndex        =   1
      Top             =   300
      Width           =   1365
      _ExtentX        =   2413
      _ExtentY        =   593
      _Version        =   393216
      Format          =   556269569
      CurrentDate     =   39795
   End
   Begin MSComCtl2.DTPicker tdate 
      Height          =   330
      Left            =   3300
      TabIndex        =   2
      Top             =   300
      Width           =   1305
      _ExtentX        =   2307
      _ExtentY        =   593
      _Version        =   393216
      Format          =   556269569
      CurrentDate     =   39795
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F4 Delete Record ...."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   228
      Index           =   2
      Left            =   108
      TabIndex        =   19
      Top             =   8604
      Width           =   1620
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   1455
      Top             =   900
      Width           =   3330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   2940
      TabIndex        =   6
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Allotment Date :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   360
      Width           =   1380
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Representative :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   228
      Left            =   9036
      TabIndex        =   4
      Top             =   252
      Visible         =   0   'False
      Width           =   1368
   End
End
Attribute VB_Name = "frmBookAllotment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub search()

On Error GoTo err:

Dim rss As New ADODB.Recordset
Screen.MousePointer = vbHourglass
vs.Cols = 6
vs.rows = 3

Dim tqty As Long
Dim Ordqty As Long
Dim Spqty As Long

k1 = 1
vs.Clear

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT * from SpAllotmentQty where len(RepName)> 0", con
If rs1.EOF = False Then
   
   
    fdate.value = rs1!fromdate
    tdate.value = rs1!todate
   
    For J = 1 To rs1.RecordCount
    
        DoEvents
        vs.TextMatrix(k1, 0) = k1
        vs.TextMatrix(k1, 1) = rs1!RepName
        vs.TextMatrix(k1, 2) = rs1!qty
        
        'If rs1!RepName = "SHIVRATAN RAWAT" Then
        'g = 0
        'End If
        
        If rss.State = 1 Then rss.close
        rss.Open "select sum(Qty) from totalSpQty_IssuedNew where RepName='" & rs1!RepName & "'and (convert(datetime,invoiceDate,103)>=convert(datetime,'" & fdate.value & "',103) and convert(datetime,invoiceDate,103)<=convert(datetime,'" & tdate.value & "',103))", con
        If Not IsNull(rss(0)) Then
           vs.TextMatrix(k1, 3) = rss(0)
        Else
           vs.TextMatrix(k1, 3) = 0
        End If
        
                
        If rss.State = 1 Then rss.close
        rss.Open "SELECT sum(QTY) FROM TotalSpReturnRepWise where agentname='" & rs1!RepName & "'  and (convert(datetime,invoiceDate,103)>=convert(datetime,'" & fdate & "',103) and convert(datetime,invoiceDate,103)<=convert(datetime,'" & tdate & "',103))", con
        If Not IsNull(rss(0)) Then
           vs.TextMatrix(k1, 4) = rss(0)
        Else
           vs.TextMatrix(k1, 4) = 0
        End If
        
        
        tqty = IIf(vs.TextMatrix(k1, 2) = "", 0, vs.TextMatrix(k1, 2))
        Ordqty = IIf(vs.TextMatrix(k1, 3) = "", 0, vs.TextMatrix(k1, 3))
        Spqty = IIf(vs.TextMatrix(k1, 4) = "", 0, vs.TextMatrix(k1, 4))
        
        
        vs.TextMatrix(k1, 5) = ((tqty + Spqty) - Ordqty)
        
        For k2 = 0 To 5
          If Val(vs.TextMatrix(k1, 5)) < 0 Then
            vs.Cell(flexcpBackColor, k1, k2) = vbGreen
            DoEvents
          End If
        Next
        
        
        rs1.MoveNext
        k1 = k1 + 1
        vs.rows = vs.rows + 1
        DoEvents
        DoEvents
    
    Next



End If

Screen.MousePointer = vbDefault

vs.FormatString = "BookCode|BookName|Allotment Qty.|TQty.Specimen|TQtyRet.Specimen|BalanceQty"
vs.ColWidth(0) = 1000
vs.ColWidth(1) = 3600
vs.ColWidth(2) = 1400
vs.ColWidth(3) = 1500
vs.ColWidth(4) = 1700
vs.ColWidth(5) = 1600

Exit Sub
err:
MsgBox "" & err.DESCRIPTION

Screen.MousePointer = vbDefault
End Sub
Private Sub Check1_UpdateFromExcel_Click()
   If Check1_UpdateFromExcel.value = 1 Then
      Frame1.Visible = True
   Else
      Frame1.Visible = False
   End If
End Sub
Private Sub cmdAddNewRep_Click()

k2 = 0

For k1 = 1 To vs.rows - 1
  If vs.TextMatrix(k1, 0) = "" Then
      k2 = k1
  End If
Next


k2 = k2 - 1

If rs1.State = 1 Then rs1.close
rs1.Open "select Rep,email from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue

While rs1.EOF = False


    If RS.State = 1 Then RS.close
    RS.Open "select RepName from SpAllotmentQty where (repname='" & rs1(0) & "')", con
    If RS.EOF = True Then
       vs.rows = vs.rows + 1
       vs.TextMatrix(k2, 0) = k2
       vs.TextMatrix(k2, 1) = rs1(0)
       vs.TextMatrix(k2, 2) = 0
       vs.TextMatrix(k2, 3) = 0
       vs.TextMatrix(k2, 4) = 0
       vs.TextMatrix(k2, 5) = 0
       
       k2 = k2 + 1
     End If
     
    
    
    rs1.MoveNext
    
 Wend

 MsgBox "Ok", vbInformation

   
End Sub
Private Sub cmdGenerateFile_Click()

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

xlSheet.Cells(1, 1).value = "BookCode"
xlSheet.Cells(1, 2).value = "BookName"
xlSheet.Cells(1, 3).value = "Sp.AllotmentQty"


K = 2

For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 0) <> "" Then
   xlSheet.Cells(K, 1).value = vs.TextMatrix(I, 0)
   xlSheet.Cells(K, 2).value = vs.TextMatrix(I, 1)
   xlSheet.Cells(K, 3).value = vs.TextMatrix(I, 2)
   K = K + 1
End If

Next

cmdGenerateFile.Enabled = False

Exit Sub
err:
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub cmdPath_Click()
Me.CommonDialog1.ShowOpen
Me.txtpath.text = Me.CommonDialog1.filename
cmdUpdateSp.Enabled = True
End Sub

Private Sub cmdUpdateSp_Click()


On Error GoTo aa:

'If cboRep.text = "" Then
'   MsgBox "Select Rep. Name ....", vbInformation
'   cboRep.SetFocus
'   Exit Sub
'End If

Dim sconn As String
Dim saveData As New ADODB.Recordset
Dim I As Integer
sFile = Me.txtpath
Screen.MousePointer = vbHourglass

'sconn = "DRIVER=Microsoft Excel Driver (*.xls,*.xlsx);" & "DBQ=" & sFile
sconn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" & "DBQ=" & sFile & ";"
If RS.State = 1 Then RS.close
RS.Open "SELECT " & _
" * FROM [Sheet4$]", sconn, adOpenDynamic, adLockReadOnly
While RS.EOF = False
con.Execute "update SpAllotmentQty set Qty=" & RS(2) & ",uid=" & UId & " where bookcode='" & RS(0) & "' and RepName='" & RS(1) & "'"
RS.MoveNext
Wend
      
Screen.MousePointer = vbDefault
MsgBox "updated data....", vbInformation
      
Exit Sub

aa:

Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION
      
 
End Sub
Private Sub CommandPrint_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim str_ As String




If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double



row_ = 1
col_ = 1

xl.Columns("A:H").ColumnWidth = 12
J = 2


For I = 0 To vs.rows - 1
    For J = 0 To vs.Cols - 1
      
        xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
       
        col_ = col_ + 1
    Next
    row_ = row_ + 1
    col_ = 1
Next

    
End Sub

Private Sub CommandReturn_Click()
Unload Me
End Sub
Private Sub Commandsave_Click()

Screen.MousePointer = vbHourglass

For J = 1 To vs.rows - 1
If vs.TextMatrix(J, 0) <> "" Then

  If RS.State = 1 Then RS.close
  RS.Open "select * from SpAllotmentQty where (repname='" & vs.TextMatrix(J, 1) & "')", con, adOpenDynamic, adLockOptimistic
  If RS.EOF = True Then
     RS.AddNew
  End If
  ''RS!Bookcode = vs.TextMatrix(J, 0)
  RS!RepName = vs.TextMatrix(J, 1)
  RS!fromdate = fdate.value
  RS!todate = tdate.value
  RS!qty = IIf(vs.TextMatrix(J, 2) = "", 0, vs.TextMatrix(J, 2))
  RS!UId = UId
  RS.update
  
End If
Next


Screen.MousePointer = vbDefault

MsgBox "Data Saved...", vbInformation

End Sub
Private Sub Form_Load()
   
Me.Width = 11535
Me.Height = 9800
Me.top = 100
Me.Left = 100
   
Dim k1 As Integer
Dim ss1 As String

ss1 = ""
vs.rows = 1

   
If rs1.State = 1 Then rs1.close
rs1.Open "SELECT * from SpAllotmentQty where len(RepName)> 0", con
If rs1.EOF = True Then
   
    If RS.State = 1 Then RS.close
    RS.Open "select Rep as Representative from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
    sss = RS.RecordCount
    cboRep.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            vs.TextMatrix(k1, 0) = k1
            vs.TextMatrix(k1, 1) = RS(0)
            
             
            vs.rows = vs.rows + 1
            k1 = k1 + 1
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If

Else

   search
   
End If


''=========================
If RS.State = 1 Then RS.close
RS.Open "select Rep as Representative from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
If Not RS.EOF Then
Do While Not RS.EOF
   If IsNull(RS(0)) = False Then
     If ss1 = "" Then
        ss1 = RS(0)
     Else
        ss1 = ss1 & "|" & RS(0)
     End If
   End If
   If Not RS.EOF Then RS.MoveNext
 Loop
End If


vs.ColComboList(1) = ss1



vs.FormatString = "BookCode|BookName|Allotment Qty.|TQty.Specimen|TQtyRet.Specimen|BalanceQty"
vs.ColWidth(0) = 1000
vs.ColWidth(1) = 3500
vs.ColWidth(2) = 1400
vs.ColWidth(3) = 1500
vs.ColWidth(4) = 1650
vs.ColWidth(5) = 1550
    
Check1_UpdateFromExcel.Visible = False
If (LCase(UserName) = "admin" Or LCase(UserName) = "dc" Or LCase(UserName) = "rishabh") Then
 Check1_UpdateFromExcel.Visible = True
End If
    
End Sub

Private Sub vs_ChangeEdit()
  vs.TextMatrix(vs.RowSel, 0) = vs.Row
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 115 Then
  
  con.Execute "delete from SpAllotmentQty where (Qty='" & vs.TextMatrix(vs.RowSel, 2) & "' and repname='" & vs.TextMatrix(vs.RowSel, 1) & "')"
  vs.RemoveItem (vs.RowSel)
  
  MsgBox "Deleted ...", vbInformation
End If

End Sub
