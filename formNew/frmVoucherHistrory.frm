VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVoucherHistory 
   Caption         =   "Voucher Recording .."
   ClientHeight    =   10116
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   18108
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVoucherHistrory.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10116
   ScaleWidth      =   18108
   Begin VB.TextBox txtVNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11628
      TabIndex        =   10
      Top             =   216
      Width           =   840
   End
   Begin VB.OptionButton Option4_All 
      Caption         =   "All"
      Height          =   264
      Left            =   10620
      TabIndex        =   9
      Top             =   252
      Value           =   -1  'True
      Width           =   876
   End
   Begin VB.OptionButton Option3_R 
      Caption         =   "R"
      Height          =   264
      Left            =   9756
      TabIndex        =   8
      Top             =   252
      Width           =   768
   End
   Begin VB.OptionButton Option2_P 
      Caption         =   "P"
      Height          =   264
      Left            =   8928
      TabIndex        =   7
      Top             =   252
      Width           =   768
   End
   Begin VB.OptionButton Option1_J 
      Caption         =   "J"
      Height          =   264
      Left            =   8028
      TabIndex        =   6
      Top             =   252
      Width           =   804
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4404
      Left            =   36
      TabIndex        =   0
      Top             =   864
      Width           =   17916
      _cx             =   31602
      _cy             =   7768
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
      FixedCols       =   1
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVoucherHistrory.frx":000C
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
   Begin VSFlex7Ctl.VSFlexGrid vs1 
      Height          =   4116
      Left            =   36
      TabIndex        =   1
      Top             =   5796
      Width           =   17988
      _cx             =   31729
      _cy             =   7260
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
      FixedCols       =   1
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVoucherHistrory.frx":00CD
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
   Begin MSComCtl2.DTPicker txtvDate 
      Height          =   336
      Left            =   5724
      TabIndex        =   4
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
      Format          =   536281089
      CurrentDate     =   39795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date :"
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
      Left            =   3924
      TabIndex        =   5
      Top             =   216
      Width           =   1956
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Voucher :-"
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
      Left            =   72
      TabIndex        =   3
      Top             =   504
      Width           =   2856
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recording Voucher :-"
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
      Left            =   72
      TabIndex        =   2
      Top             =   5436
      Width           =   2856
   End
End
Attribute VB_Name = "frmVoucherHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
   Unload Me
End If

End Sub

Private Sub Form_Load()

Me.top = 100
Me.Left = 100

Me.Width = 18500
Me.Height = 11000

txtvDate.value = Format(Date, "dd/MM/yyyy")

search ""

End Sub
Sub search(vno1 As String)

   Dim vtype As String

   If Option1_J.value = True Then
      vtype = "J"
   ElseIf Option2_P.value = True Then
      vtype = "P"
   ElseIf Option3_R.value = True Then
      vtype = "R"
   Else
      vtype = "All"
   End If
 
 
   Dim r1 As Integer
   vs.Clear
   
   vs1.Clear
   
   
   r1 = 1
   
   vs.rows = 2
   vs.Cols = 11
 
   If rs1.State = 1 Then rs1.close
   
   If vtype = "All" Then
     If vno1 = "" Then
        rs1.Open "Select * from vouchers where convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & txtvDate.value & "',103) order by vsno", con, adOpenDynamic, adLockOptimistic
     Else
        rs1.Open "Select * from vouchers where VoucherNumber=" & vno1 & " and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & txtvDate.value & "',103) order by vsno", con, adOpenDynamic, adLockOptimistic
     End If
   Else
     If vno1 = "" Then
      rs1.Open "Select * from vouchers where VoucherType='" + vtype + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & txtvDate.value & "',103) order by vsno", con, adOpenDynamic, adLockOptimistic
     Else
      rs1.Open "Select * from vouchers where VoucherNumber=" & vno1 & " and VoucherType='" + vtype + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & txtvDate.value & "',103) order by vsno", con, adOpenDynamic, adLockOptimistic
     
     End If
   End If
   
   While rs1.EOF = False
   
   
   vs.TextMatrix(r1, 0) = rs1!VoucherType
   vs.TextMatrix(r1, 1) = rs1!voucherDATE
   
   vs.TextMatrix(r1, 2) = rs1!VOUCHERNUMBER
   vs.TextMatrix(r1, 3) = rs1!Genledger
   vs.TextMatrix(r1, 4) = rs1!subledger
   vs.TextMatrix(r1, 5) = rs1!amount
   
   vs.TextMatrix(r1, 6) = rs1!DebitorCredit
   vs.TextMatrix(r1, 7) = rs1!CBND
   vs.TextMatrix(r1, 8) = rs1!DESCRIPTION & ""
   
   vs.TextMatrix(r1, 9) = rs1!UserName
   
   vs.TextMatrix(r1, 10) = rs1!dates
   
      
   
   vs.rows = vs.rows + 1
   r1 = r1 + 1
   rs1.MoveNext
   Wend
   
   
''==========================================================================================================
   
   r1 = 1
 
   vs1.Cols = 12
   vs1.rows = 2
 
   If rs1.State = 1 Then rs1.close
   
   If vtype = "All" Then
     
     If vno1 = "" Then
        rs1.Open "Select * from vouchers_bk where convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & txtvDate.value & "',103) order by DATES,VSNO", con, adOpenDynamic, adLockOptimistic
     Else
        rs1.Open "Select * from vouchers_bk where VoucherNumber=" & vno1 & " and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & txtvDate.value & "',103) order by DATES,VSNO", con, adOpenDynamic, adLockOptimistic
     
     End If
   Else
     If vno1 = "" Then
         rs1.Open "Select * from vouchers_bk where VoucherType='" + vtype + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & txtvDate.value & "',103) order by DATES,VSNO", con, adOpenDynamic, adLockOptimistic
     Else
         rs1.Open "Select * from vouchers_bk where VoucherNumber=" & vno1 & " and VoucherType='" + vtype + "' and convert(smalldatetime,voucherdate,103)= convert(smalldatetime,'" & txtvDate.value & "',103) order by DATES,VSNO", con, adOpenDynamic, adLockOptimistic
     End If
   End If
   
   While rs1.EOF = False
   
   
   vs1.TextMatrix(r1, 0) = rs1!VoucherType & " : " & rs1!paytype
   vs1.TextMatrix(r1, 1) = rs1!voucherDATE
   
   vs1.TextMatrix(r1, 2) = rs1!VOUCHERNUMBER
   vs1.TextMatrix(r1, 3) = rs1!Genledger
   vs1.TextMatrix(r1, 4) = rs1!subledger
   vs1.TextMatrix(r1, 5) = rs1!amount
   
   vs1.TextMatrix(r1, 6) = rs1!DebitorCredit
   vs1.TextMatrix(r1, 7) = rs1!CBND
   vs1.TextMatrix(r1, 8) = rs1!DESCRIPTION & ""
   
   vs1.TextMatrix(r1, 9) = rs1!UserName
   
   vs1.TextMatrix(r1, 10) = rs1!dates
   
   
      
   
   vs1.rows = vs1.rows + 1
   r1 = r1 + 1
   
   rs1.MoveNext
   Wend


''=======================================================================================================
On Error Resume Next
For k1 = 1 To vs.rows - 1
   If vs.TextMatrix(k1, 5) <> vs1.TextMatrix(k1, 5) Then
      For kk1 = 0 To 8
                vs1.Cell(flexcpBackColor, k1, kk1) = vbGreen
                DoEvents
                vs.Cell(flexcpBackColor, k1, kk1) = vbGreen
                DoEvents

      Next
   End If
Next
''=======================================================================================================
   
vs.WordWrap = True
   
vs.FormatString = "Vtye|VDate|VNo.|GLedger|Subledger|Amount|Dr/Cr|ChequeDetails|DESCRIPTION|UName|Dates"
vs.ColWidth(0) = 600
vs.ColWidth(1) = 1400
vs.ColWidth(2) = 1000
vs.ColWidth(3) = 3500
vs.ColWidth(4) = 4000
vs.ColWidth(5) = 1500
vs.ColWidth(6) = 900
vs.ColWidth(7) = 2500
vs.ColWidth(8) = 4500
vs.ColWidth(9) = 1000
vs.ColWidth(10) = 2500

vs1.WordWrap = True
   
vs1.FormatString = "Vtye|VDate|VNo.|GLedger|Subledger|Amount|Dr/Cr|ChequeDetails|DESCRIPTION|UName|Dates"
vs1.ColWidth(0) = 1500
vs1.ColWidth(1) = 1400
vs1.ColWidth(2) = 1000
vs1.ColWidth(3) = 3500
vs1.ColWidth(4) = 4000
vs1.ColWidth(5) = 1500
vs1.ColWidth(6) = 900
vs1.ColWidth(7) = 2500
vs1.ColWidth(8) = 4500
vs1.ColWidth(9) = 1000
vs1.ColWidth(10) = 2500




End Sub

Private Sub Option1_J_Click()

If txtVNo.text = "" Then
  search ""
Else
  search txtVNo.text
End If

End Sub

Private Sub Option2_P_Click()
If txtVNo.text = "" Then
  search ""
Else
  search txtVNo.text
End If
End Sub

Private Sub Option3_R_Click()
If txtVNo.text = "" Then
  search ""
Else
  search txtVNo.text
End If
End Sub

Private Sub Option4_All_Click()
If txtVNo.text = "" Then
  search ""
Else
  search txtVNo.text
End If
End Sub

Private Sub txtVNo_Change()
If txtVNo.text = "" Then
  search ""
Else
  search txtVNo.text
End If
End Sub
