VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAuditTrailLog 
   Caption         =   "Audit Trail Log"
   ClientHeight    =   10200
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   18072
   Icon            =   "frmAuditTrailLog.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   18072
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   624
      Left            =   15048
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   72
      Width           =   1248
   End
   Begin VB.OptionButton Option3_Edit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11304
      TabIndex        =   9
      Top             =   180
      Value           =   -1  'True
      Width           =   1164
   End
   Begin VB.OptionButton Option2_del 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12528
      TabIndex        =   8
      Top             =   180
      Width           =   1236
   End
   Begin VB.OptionButton Option1_Insert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10044
      TabIndex        =   7
      Top             =   180
      Width           =   1200
   End
   Begin VB.OptionButton Option4_All 
      Caption         =   "ALL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13824
      TabIndex        =   2
      Top             =   180
      Width           =   1056
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
      Left            =   6264
      TabIndex        =   1
      Top             =   144
      Width           =   3480
   End
   Begin VB.CommandButton CommandReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   16344
      Picture         =   "frmAuditTrailLog.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   72
      Width           =   1116
   End
   Begin MSComCtl2.DTPicker fDate_ 
      Height          =   336
      Left            =   144
      TabIndex        =   3
      Top             =   144
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
      Format          =   141361153
      CurrentDate     =   39795
   End
   Begin MSComCtl2.DTPicker tdate_ 
      Height          =   336
      Left            =   2412
      TabIndex        =   4
      Top             =   144
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
      Format          =   141361153
      CurrentDate     =   39795
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   9048
      Left            =   108
      TabIndex        =   10
      Top             =   720
      Width           =   17700
      _cx             =   31221
      _cy             =   15960
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761992
      ForeColorSel    =   0
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   200
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   450
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAuditTrailLog.frx":0BF0
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
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :-"
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
      Left            =   216
      TabIndex        =   11
      Top             =   9828
      Width           =   2100
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
      Left            =   4284
      TabIndex        =   6
      Top             =   180
      Width           =   2352
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
      Left            =   1908
      TabIndex        =   5
      Top             =   180
      Width           =   516
   End
End
Attribute VB_Name = "frmAuditTrailLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub search()

    
   Dim value_ As String
   Dim actionType_ As String
   
   Dim rs_fill As ADODB.Recordset
   
   value_ = "(invoicedate>=convert(smalldatetime,'" & fDate_.value & "',103) and invoicedate<=convert(smalldatetime,'" & tdate_.value & "',103))"

   

   Dim vtype1 As String
   actionType_ = ""
   
   If (Option1_Insert.value = True) Then
      actionType_ = "Insert"
   ElseIf (Option3_Edit.value = True) Then
      actionType_ = "Edit"
   ElseIf (Option2_del.value = True) Then
      actionType_ = "Delete"
   ElseIf (Option4_All.value = True) Then
      actionType_ = "ALL"
   End If
   
   
   
 
If (cbovtype.text = "Payment Voucher" Or cbovtype.text = "Receipt Voucher" Or cbovtype.text = "Journal Voucher") Then
   vtype1 = Left(cbovtype.text, 1)
Else
   vtype1 = cbovtype.text
End If
        
        
      
        
If actionType_ = "ALL" Then

  value_ = " (VoucherDate>=convert(smalldatetime,'" & fDate_.value & "',103) and VoucherDate<=convert(smalldatetime,'" & tdate_.value & "',103)) and vouchertype='" & vtype1 & "'"

Else

   value_ = " (VoucherDate>=convert(smalldatetime,'" & fDate_.value & "',103) and VoucherDate<=convert(smalldatetime,'" & tdate_.value & "',103)) and ActionType='" & actionType_ & "' and vouchertype='" & vtype1 & "'"

End If



str_ = "Select VoucherID as VNo,VoucherType as VType,ActionType,FORMAT(VoucherDate,'dd/MM/yyyy') as VDate," & _
      "VoucherNumber as VN,[Description],FORMAT(Dates,'dd/MM/yyyy') as ActionDate,Amount,ReasionForEdit,UserName,Id from AuditTrail_Log   where  " + value_ + "" & _
      "order  BY ID"






'    End If
        

 
 
   Dim r1 As Integer
   'vs.Clear

 
   
   
   r1 = 1
   
   vs.Clear
   
   vs.rows = 2
   vs.Cols = 9
   
   action_v = ""
'   Dim k1 As Integer
'
'   k1 = 1
'
'   If rs1.State = 1 Then rs1.close
'
'
   Set rs_fill = New ADODB.Recordset
   
   rs_fill.Open str_, con, adOpenDynamic, adLockOptimistic
'
'
'   While rs1.EOF = False
'
'   vs.TextMatrix(k1, 0) = rs1!vno
'   vs.TextMatrix(k1, 1) = rs1!vtype
'   vs.TextMatrix(k1, 2) = rs1!ActionType
'
'   vs.TextMatrix(k1, 3) = rs1!vdate
'   vs.TextMatrix(k1, 4) = rs1!vn
'   vs.TextMatrix(k1, 5) = rs1!ActionDate
'
'   vs.TextMatrix(k1, 6) = rs1!Amount
'   vs.TextMatrix(k1, 7) = rs1!ReasionForEdit & ""
'   vs.TextMatrix(k1, 8) = rs1!Id
'
'
'
'   k1 = k1 + 1
'   vs.rows = vs.rows + 1
'
'   rs1.MoveNext
'
'   Wend
   
   Set vs.DataSource = rs_fill
   
''==========================================================================================================
   
lblTotal.Caption = "Total : " & vs.rows - 1

vs.WordWrap = True

vs.FormatString = "VNO|VoucherType|ActionType|VDate|VN|Description|ActionDate|Amount|Reason For Edit|UName|Id"
vs.ColWidth(0) = 800
vs.ColWidth(1) = 1500
vs.ColWidth(2) = 1100
vs.ColWidth(3) = 1500
vs.ColWidth(4) = 800
vs.ColWidth(5) = 3800
vs.ColWidth(6) = 1500
vs.ColWidth(7) = 1500
vs.ColWidth(8) = 3200
vs.ColWidth(9) = 1000
vs.ColWidth(10) = 600





End Sub

Private Sub cbovtype_Click()
search
End Sub

Private Sub cmdRepQty_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset



If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

    xlSheet.Range("D1").HorizontalAlignment = xlLeft
    
 

    row_ = 1
    col_ = 1
   
    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    xlSheet.Cells(1, 1).value = "Audit Trail Log "
    
    For I = 0 To vs.rows - 1
        For J = 0 To vs.Cols - 1
        
            If IsDate(vs.TextMatrix(I, J)) Then
               strDate = FormatDateTime(vs.TextMatrix(I, J), vbShortDate)
               xlSheet.Cells(row_, col_).value = strDate
  
            Else
               xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
            End If
          
            
            col_ = col_ + 1
            
        Next
        row_ = row_ + 1
        col_ = 1
    Next
    
    Screen.MousePointer = vbDefault



End Sub

Private Sub CommandReturn_Click()
Unload Me
End Sub

Private Sub Form_Load()


cbovtype.Clear

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT VType_Det FROM CheckedBy group by VType_Det", con
While rs1.EOF = False

cbovtype.AddItem rs1(0)
rs1.MoveNext

Wend


Me.top = 100
Me.Left = 100

Me.Width = 19000
Me.Height = 10900

fDate_.value = Format(from_date, "dd/MM/yyyy")
tdate_.value = Format(to_date, "dd/MM/yyyy")

End Sub


Private Sub Option1_Insert_Click()
search
End Sub

Private Sub Option2_del_Click()
search
End Sub

Private Sub Option3_Edit_Click()
search
End Sub

Private Sub Option4_All_Click()
search
End Sub
