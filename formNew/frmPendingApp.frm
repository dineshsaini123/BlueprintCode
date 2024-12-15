VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmPendingApp 
   Caption         =   "Pending Approval"
   ClientHeight    =   8328
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   14928
   Icon            =   "frmPendingApp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8328
   ScaleWidth      =   14928
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboembp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      ItemData        =   "frmPendingApp.frx":000C
      Left            =   1512
      List            =   "frmPendingApp.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   864
      Width           =   1236
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   10530
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   225
      Width           =   1104
   End
   Begin VB.CommandButton CommandReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12816
      Picture         =   "frmPendingApp.frx":0022
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   225
      Width           =   1056
   End
   Begin VB.CommandButton CommandPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11664
      Picture         =   "frmPendingApp.frx":0C06
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   225
      Width           =   1128
   End
   Begin VB.TextBox txtSchoolName 
      Height          =   315
      Left            =   4140
      MaxLength       =   150
      TabIndex        =   5
      Top             =   495
      Width           =   5325
   End
   Begin VB.TextBox txtscid 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9510
      TabIndex        =   4
      Top             =   495
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   3975
      Begin VB.OptionButton Option1_All 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2565
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1_school 
         Caption         =   "School"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   1185
      End
      Begin VB.OptionButton Option2_Party 
         Caption         =   "Party"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   270
         Width           =   960
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6648
      Left            =   48
      TabIndex        =   0
      Top             =   1284
      Width           =   14724
      _cx             =   25971
      _cy             =   11726
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   7917545
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPendingApp.frx":17EA
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BP/EM :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   504
      TabIndex        =   13
      Top             =   936
      Width           =   936
   End
   Begin VB.Label lblrow 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   1
      Left            =   96
      TabIndex        =   11
      Top             =   8088
      Width           =   2916
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "School Name  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4185
      TabIndex        =   6
      Top             =   180
      Width           =   2910
   End
End
Attribute VB_Name = "frmPendingApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dt_str, fdate1, tdate2 As String
Private Sub cmdView_Click()
  
  Dim st_ As String
  Dim k1 As Integer
  k1 = 0
  
  
  Screen.MousePointer = vbHourglass
  
  If RS.State = 1 Then RS.close
  RS.Open "select fromDate,toDate,NotCreated,current_next from turnOverDis order by fyear", CCON
  If RS.EOF = False Then
  If RS!NotCreated = "y" Then
   dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromdate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!todate & "',103))"
   fdate1 = RS!fromdate
   tdate2 = RS!todate
   RS.MoveNext
  End If
  
  If (RS!NotCreated = "y" And RS!current_next = "next") Then
     
     tdate2 = RS!todate
     
     dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & fdate1 & "',103) and INVOICEDATE <= convert(smalldatetime,'" & tdate2 & "',103))"
     
  End If
  
  
  End If
  
  
  
  'If frmApproval.Option1_school.value = True Then
     con.Execute "exec updateInvoiceB_withAppForm '" & fdate1 & "','" & tdate2 & "','School'"
  'Else
     DoEvents
     DoEvents
     DoEvents
     DoEvents
     con.Execute "exec updateInvoiceB_withAppForm '" & fdate1 & "','" & tdate2 & "','Party'"
  'End If
  
  DoEvents
  DoEvents
  DoEvents
  DoEvents
  
  vs.rows = 2
  
  Dim s_ As String
  
  If cboembp.text = "EM" Then
     s_ = "CHARINDEX('(EM)', SUBLEDGER)>0"
  Else
     s_ = "CHARINDEX('(EM)', SUBLEDGER)=0"
  End If
  
  
  If Option1_All.value = True Then
     st_ = "SELECT SUBLEDGER,ScName,ScID,SerName,INVOICEDATE,discount,InvoiceNo from useForApprovalQry where (appno='' and " & dt_str & " and " & s_ & ") order by SUBLEDGER,ScName,ScID"
  ElseIf Option1_school.value = True Then
     st_ = "SELECT SUBLEDGER,ScName,ScID,SerName,INVOICEDATE,discount,InvoiceNo from useForApprovalQry where (ScID ='" & txtScId.text & "' and appno='' and " & dt_str & " and " & s_ & ") order by SUBLEDGER,ScName,ScID"
  ElseIf Option2_Party.value = True Then
     st_ = "SELECT SUBLEDGER,ScName,ScID,SerName,INVOICEDATE,discount,InvoiceNo from useForApprovalQry where (pcode ='" & txtScId.text & "' and appno='' and " & dt_str & " and " & s_ & ") order by SUBLEDGER,ScName,ScID"
  End If
  
  ''If (cboParty.Text = "O2020 ONLINE SALES (PAY U)" Or cboParty.Text = "O2020 ONLINE SALES (PAY U)" Or cboParty.Text = "A2020 AMAZON.IN") Then
  
  If rs1.State = 1 Then rs1.close
  rs1.Open st_, con
  
  For K = 1 To rs1.RecordCount
  
   DoEvents
   ''DoEvents
   
   vs.TextMatrix(K, 0) = rs1!subledger & ""
   vs.TextMatrix(K, 1) = rs1!scname & ""
   vs.TextMatrix(K, 2) = rs1!scid & ""
   vs.TextMatrix(K, 3) = rs1!sername & ""
   vs.TextMatrix(K, 4) = rs1!invoiceDate & ""
   vs.TextMatrix(K, 5) = rs1!invoiceNo & ""
   vs.TextMatrix(K, 6) = rs1!discount & ""
         
   

'   If k1 = 0 Then
'    vs.Cell(flexcpBackColor, K, 0) = vbWhite
'    vs.Cell(flexcpBackColor, K, 1) = vbWhite
'    vs.Cell(flexcpBackColor, K, 2) = vbWhite
'    vs.Cell(flexcpBackColor, K, 3) = vbWhite
'    vs.Cell(flexcpBackColor, K, 4) = vbWhite
'    vs.Cell(flexcpBackColor, K, 5) = vbWhite
'  Else
'    vs.Cell(flexcpBackColor, K, 0) = &HC0FFFF
'    vs.Cell(flexcpBackColor, K, 1) = &HC0FFFF
'    vs.Cell(flexcpBackColor, K, 2) = &HC0FFFF
'    vs.Cell(flexcpBackColor, K, 3) = &HC0FFFF
'    vs.Cell(flexcpBackColor, K, 4) = &HC0FFFF
'    vs.Cell(flexcpBackColor, K, 5) = &HC0FFFF
'  End If

   
   
   vs.rows = vs.rows + 1
   rs1.MoveNext
   
   
'   If rs1.EOF = False Then
'   If vs.TextMatrix(K, 0) <> rs1!subledger Then
'       If k1 = 0 Then
'          k1 = 1
'       Else
'          k1 = 0
'       End If
'   End If
'   End If
'   DoEvents
   
   
  Next
  
  
  vs.FormatString = "SUBLEDGER|SCNAME|SCID|SERNAME|INVOICEDATE|INVOICE_NO|DISCOUNT"
  vs.ColWidth(0) = 4200
  vs.ColWidth(1) = 4000
  vs.ColWidth(2) = 1000
  vs.ColWidth(3) = 1600
  vs.ColWidth(4) = 1100
  vs.ColWidth(4) = 1000
  vs.ColWidth(5) = 1000
  
  lblrow(1).Caption = "Total Record : " & vs.rows - 1
  
  Screen.MousePointer = vbDefault
  
  
  
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
       
        If col_ = 5 Then
        xlSheet.Cells(row_, col_).NumberFormat = "dd/mm/yyyy"
        End If
       
        col_ = col_ + 1
    Next
    row_ = row_ + 1
    col_ = 1
Next

    
End Sub

Private Sub CommandReturn_Click()
Unload Me
End Sub
Private Sub Form_Load()

Me.top = 500
Me.Left = 500
Me.Width = 14990

If RS.State = 1 Then RS.close
RS.Open "select fromDate,toDate,NotCreated from turnOverDis where fyear='" & session & "'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
   dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromdate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!todate & "',103))"
   fdate1 = RS!fromdate
   tdate2 = RS!todate
  End If
End If


cboembp.ListIndex = 0

End Sub
Private Sub txtSchoolName_GotFocus()
   
If PopUpValue1 <> "" Then

If Option1_school.value = True Then
   txtSchoolName.text = PopUpValue1
   txtScId.text = PopUpValue2
Else
   txtSchoolName.text = PopUpValue1
   txtScId.text = PopUpValue3
End If
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
Else

End If
   
End Sub

Private Sub txtSchoolName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    If Option1_school.value = True Then
        Screen.MousePointer = vbHourglass
        searchType = "party"
        value = "SELECT ScName,ScID FROM invoicea where " & dt_str & " group by ScName,ScID"
        popuplistModel10 value, con
        set_focus = True
        Screen.MousePointer = vbDefault
    Else
       Screen.MousePointer = vbHourglass
       searchType = "party"
       value = "SELECT  DESCFORINVOICE as PartyName,address3 as City,Code FROM SLEDGER where len(DESCFORINVOICE)>0 order by DESCFORINVOICE"
       popuplistModel10 value, con
       set_focus = True
       Screen.MousePointer = vbDefault
    End If
End If


If KeyCode = 13 Then
   cmdView_Click
End If

End Sub
