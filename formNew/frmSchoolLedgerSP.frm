VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSchoolLedgerSP 
   Caption         =   "School Ledger For Specimen Qty."
   ClientHeight    =   9096
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   12144
   Icon            =   "frmSchoolLedgerSP.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9096
   ScaleWidth      =   12144
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FAEFC9&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2385
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   825
      Width           =   990
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FAEFC9&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   825
      Width           =   990
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   12720
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "SUNDRY DEBTORS"
      Top             =   465
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3945
      TabIndex        =   4
      Top             =   1125
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Commandreturn 
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
      Height          =   300
      Left            =   3945
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.ComboBox Combosubledger 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   312
      Left            =   1350
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   345
      Width           =   7764
   End
   Begin VB.TextBox Alpha 
      Height          =   315
      Left            =   12780
      MaxLength       =   1
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtscid 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Top             =   345
      Width           =   984
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   90
      Top             =   8910
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox date1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   5940
      TabIndex        =   8
      Top             =   1125
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1969
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox date2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   5445
      TabIndex        =   9
      Top             =   1125
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2032
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7248
      Left            =   48
      TabIndex        =   10
      Top             =   1536
      Width           =   12000
      _cx             =   21167
      _cy             =   12785
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16251308
      ForeColor       =   16711680
      BackColorFixed  =   16251308
      ForeColorFixed  =   255
      BackColorSel    =   16448755
      ForeColorSel    =   16744448
      BackColorBkg    =   16251308
      BackColorAlternate=   16251308
      GridColor       =   255
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
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   380
      RowHeightMax    =   0
      ColWidthMin     =   400
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 For Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1350
      TabIndex        =   17
      Top             =   45
      Width           =   2715
   End
   Begin VB.Label lblDrTotal 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   6360
      TabIndex        =   16
      Top             =   8835
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblCrTotal 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   8985
      TabIndex        =   15
      Top             =   8835
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblCr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   312
      Left            =   12240
      TabIndex        =   14
      Top             =   5688
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label lblDr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   312
      Left            =   12168
      TabIndex        =   13
      Top             =   6048
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "School Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   12
      Top             =   405
      Width           =   1290
   End
   Begin VB.Label Label3 
      Caption         =   " - To - "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6525
      TabIndex        =   11
      Top             =   1125
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "frmSchoolLedgerSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub setWidth()
    
    vs.Clear
    vs.Cols = 4
    
    vs.FormatString = "^Challan No|^Challan Dates|Description|Party Name|<Remarks"
    vs.ColWidth(0) = 1000
    vs.ColWidth(1) = 1250
    vs.ColWidth(2) = 2200
    vs.ColWidth(3) = 2500
    vs.ColWidth(4) = 2500
    
   DoEvents

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Combosubledger_GotFocus()
If PopUpValue1 <> "" Then
   Dim k1 As Integer
   Dim fdate, tdate
   
   Combosubledger.Text = PopUpValue1      '& ", " & PopUpValue3
   txtScId.Text = PopUpValue2
   
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""
   
   
   lblCrTotal.Caption = 0
   '==========================================
   setWidth
   k1 = 1
   '==========================================
   If RS.State = 1 Then RS.close
    RS.Open "select fromDate,toDate,NotCreated from turnOverDis order by Current_Next", CCON
    If RS.EOF = False Then
       fdate = RS!FromDate
       tdate = RS!toDate
       RS.MoveNext
       If RS!NotCreated = "y" Then
       tdate = RS!toDate
       End If
       
       
    End If
   
   vs.rows = 1
   
   If RS.State = 1 Then RS.close
   ''Set RS = con.Execute("exec Specimen_NetQty '" & txtScId & "','" & fdate & "','" & tdate & "'")
   Set RS = con.Execute("exec Specimen_NetQty '" & txtScId & "','" & from_date & "','" & to_date & "'")
   '==========================================
   'RS.Open "SELECT Invoiceno,INVOICEDATE,SUBLEDGER,NETAMOUNT FROM INVOICEA where ScID='" & txtscid & "'  order by Invoiceno", con
   For I = 1 To RS.RecordCount
   vs.rows = vs.rows + 1
   vs.TextMatrix(k1, 0) = RS!invoiceNo
   'vs.TextMatrix(k1, 1) = RS!fyear
   vs.TextMatrix(k1, 1) = RS!invoiceDate
   vs.TextMatrix(k1, 2) = "Specimen"
   vs.TextMatrix(k1, 3) = RS!agentname
   vs.TextMatrix(k1, 4) = RS!remarks
   k1 = k1 + 1
   'lblCrTotal.Caption = Val(lblCrTotal.Caption) + RS!netamount
   
   RS.MoveNext
   Next
   
  
   PopUpValue1 = ""
   PopUpValue2 = ""
   
End If
End Sub

Private Sub Combosubledger_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 113 Then

    searchType = "party"
    value = "SELECT ScName,ScID FROM INVOICEA_sp where len(ScName)>0 group by ScName,ScID"
     popuplist_client value, con
    set_focus = True

End If

'If KeyCode = 113 Then
'
'   Screen.MousePointer = vbHourglass
'   tblNo = 9
'   frmSearchItem.Show
'   Screen.MousePointer = vbDefault
'
'End If


End Sub

Private Sub Form_Load()
Me.Top = 100
Me.Left = 100

Me.Width = 12300
Me.Height = 9750



''vsIni

kk = 1

'dateAson.value = Date




'FromDate.value = Date
'toDate.value = Date
'from_date = FromDate.value


'setwidth

'cboop.ListIndex = 0



If RS.State = 1 Then RS.close
RS.Open "select yarfrom,yarto from setup1 where " & stringyear & "", con
If RS.EOF = False Then
   'FromDate.value = RS.Fields(0).value
   
   If (DateValue(RS!yarfrom) <= DateValue(Date) And DateValue(RS!yarto) >= DateValue(Date)) Then
      'RecDates.value = Date
   Else
      'RecDates.value = RS.Fields(1).value
   End If
   
End If

Me.Top = 50
Me.Left = 50





If RS.State = 1 Then RS.close
RS.Open "select * from setup1 where " & stringyear & "", con
If RS.EOF = False Then
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
End If



bb1 = False


BackColorFrom Me

Screen.MousePointer = vbDefault

End Sub
Private Sub vs_DblClick()



If vs.TextMatrix(vs.RowSel, 1) <> "" Then
'  If session = vs.TextMatrix(vs.RowSel, 1) Then
    inviceNo = vs.TextMatrix(vs.RowSel, 0)
    frmBookIssueSp.Show
'  End If
End If


End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   'If vs.TextMatrix(vs.RowSel, 3) = "Specimen" Then
   'If session = vs.TextMatrix(vs.RowSel, 1) Then
    inviceNo = vs.TextMatrix(vs.RowSel, 0)
    frmBookIssueSp.Show
   'End If
'End If
End If

End Sub
