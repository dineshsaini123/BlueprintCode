VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmReel_ReamRpt 
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   17685
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtBal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   16275
      TabIndex        =   23
      Top             =   2400
      Width           =   915
   End
   Begin VB.TextBox txtIssue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   16275
      TabIndex        =   22
      Top             =   1950
      Width           =   915
   End
   Begin VB.TextBox txtRec 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   16275
      TabIndex        =   21
      Top             =   1500
      Width           =   915
   End
   Begin Crystal.CrystalReport cr 
      Left            =   16350
      Top             =   4650
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   12525
      TabIndex        =   19
      Top             =   450
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   14850
      TabIndex        =   17
      Top             =   450
      Width           =   1065
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   465
      Left            =   13725
      TabIndex        =   16
      Top             =   450
      Width           =   1065
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   465
      Left            =   11325
      TabIndex        =   15
      Top             =   450
      Width           =   1065
   End
   Begin VB.ComboBox cboGSM 
      Height          =   315
      Left            =   3975
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   525
      Width           =   1290
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   5700
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   525
      Width           =   2190
   End
   Begin VB.ComboBox cboMillName 
      Height          =   315
      Left            =   8775
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   525
      Width           =   2190
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   11100
      TabIndex        =   0
      Top             =   375
      Visible         =   0   'False
      Width           =   90
      Begin VB.OptionButton RealOption 
         Caption         =   "Reel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   2
         Top             =   150
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton ReamOption 
         Caption         =   "Ream"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   975
         TabIndex        =   1
         Top             =   150
         Width           =   990
      End
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   315
      Left            =   675
      TabIndex        =   6
      Top             =   525
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   315
      Left            =   2175
      TabIndex        =   11
      Top             =   525
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6975
      Left            =   75
      TabIndex        =   18
      Top             =   1275
      Width           =   14685
      _cx             =   25903
      _cy             =   12303
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483635
      BackColorSel    =   13888387
      ForeColorSel    =   16711680
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   7
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmReel_ReamRpt.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.Label Label8 
      Caption         =   "Total Balance :"
      Height          =   315
      Left            =   15075
      TabIndex        =   25
      Top             =   2400
      Width           =   1290
   End
   Begin VB.Label Label6 
      Caption         =   "Total Issued :"
      Height          =   315
      Left            =   15075
      TabIndex        =   24
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Total Received :"
      Height          =   315
      Left            =   15075
      TabIndex        =   20
      Top             =   1575
      Width           =   1290
   End
   Begin VB.Label line1 
      BackColor       =   &H00FFC0C0&
      Height          =   165
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   1050
      Width           =   20490
   End
   Begin VB.Label line1 
      BackColor       =   &H00FFC0C0&
      Height          =   165
      Index           =   0
      Left            =   -150
      TabIndex        =   13
      Top             =   0
      Width           =   20640
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   240
      Left            =   1875
      TabIndex        =   12
      Top             =   525
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "GSM :"
      Height          =   240
      Left            =   3450
      TabIndex        =   10
      Top             =   525
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Size :"
      Height          =   240
      Left            =   5250
      TabIndex        =   9
      Top             =   525
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Mill Name  :"
      Height          =   240
      Left            =   7950
      TabIndex        =   8
      Top             =   525
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "From"
      Height          =   240
      Left            =   225
      TabIndex        =   7
      Top             =   525
      Width           =   690
   End
End
Attribute VB_Name = "frmReel_ReamRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
 
 
 If MsgBox("Want To Print ?", vbQuestion + vbYesNo) = vbYes Then
    
  CON.BeginTrans
  CON.Execute "delete from ReelRpt"
  For i = 1 To vs.Rows - 1
  
  If vs.TextMatrix(i, 1) <> "" Then
    CON.Execute "insert into ReelRpt([ReelNo],[RecDate],[Weight],[GSM],[Size],[Mill],[IssueDate],[Status]) " & _
    "values('" & vs.TextMatrix(i, 1) & "','" & Format(vs.TextMatrix(i, 2), "MM/dd/yyyy") & "','" & vs.TextMatrix(i, 3) & "','" & vs.TextMatrix(i, 4) & "'" & _
    ",'" & vs.TextMatrix(i, 5) & "','" & vs.TextMatrix(i, 6) & "','" & Format(vs.TextMatrix(i, 7), "MM/dd/yyyy") & "','" & vs.TextMatrix(i, 8) & "')"
  End If
  
  Next
    
  CON.CommitTrans
 
 End If
 
 
 '-----------------------------------------------
 
    CR.Reset
    CR.Connect = constr
    CR.ReportFileName = strrptpath & "\reports\ReelRpt.rpt"
    CR.WindowShowPrintBtn = True
    CR.WindowShowPrintSetupBtn = True
    CR.Formulas(0) = "fromdate='" & date1.Text & "'"
    CR.Formulas(1) = "todate='" & date2.Text & "'"
    CR.Formulas(2) = "bal=" & txtBal & ""
    CR.WindowState = crptMaximized
    CR.Action = 1

 '-------------------------------------------------
 

End Sub
Private Sub cmdSave_Click()
Dim dd

If MsgBox("Want To Update ?", vbInformation + vbYesNo) = vbYes Then
For i = 1 To vs.Rows - 1
If vs.TextMatrix(i, 1) <> "" Then
   
If vs.TextMatrix(i, 8) = "" Then
 var = 0
 dd = "01/10/1900"
Else
 If vs.TextMatrix(i, 8) = True Then
   var = 1
   dd = Format(vs.TextMatrix(i, 7), "MM/dd/yyyy")
End If

End If
   

If IsDate(vs.TextMatrix(i, 7)) = True Then
CON.Execute "update [Reel_ReamDetails] set DateofIssue='" & dd & "',[Reel_Issued]='" & var & "' where [ReelNo]='" & vs.TextMatrix(i, 1) & "' and " & stringyear
Else
dd = Null
var = ""
CON.Execute "update [Reel_ReamDetails] set DateofIssue='" & dd & "',[Reel_Issued]='" & var & "' where [ReelNo]='" & vs.TextMatrix(i, 1) & "' and " & stringyear
End If

End If
Next


End If
End Sub

Private Sub cmdView_Click()

Dim rss As New ADODB.Recordset
Dim b1 As String
Dim s As String

Dim rec, Issue As Integer


s = ""
b1 = ""


If s = "" Then
   s = "(convert(smalldatetime,Dates,103)>=convert(smalldatetime,'" & Trim(date1.Text) & "',103) AND convert(smalldatetime,Dates,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103))"
End If

  
  If cboGSM.Text <> "" Then
     s = s & " and " & "(gsm='" & cboGSM.Text & "'"
     b1 = "y"
  End If
   
  If cboSize.Text <> "" Then
     
     If b1 = "" Then
        s = s & " and " & "(Size='" & cboSize.Text & "'"
     
     Else
        s = s & " and " & "Size='" & cboSize.Text & "'"
     End If
     
     
     b1 = "y"
  End If
   
   
  If cboMillName.Text <> "" Then
     
     If b1 = "" Then
        s = s & " and " & "(Mill='" & cboMillName.Text & "'"
     
     Else
     
       s = s & " and " & "Mill='" & cboMillName.Text & "'"
       
     End If
     
     
     b1 = "y"
  
  End If
   
  If b1 = "y" Then
     s = s & ")"
  End If



vs.Clear



'======================================================================================================
'======================================================================================================
vs.Cols = 9

If rss.State = 1 Then rss.Close
rss.Open "select ReelNo,Weight,GSM,Size,Mill,[DateofIssue],[Reel_Issued] from Reel_ReamDetails WHERE " & s & " and " & stringyear & "", CON, adOpenKeyset, adLockReadOnly
ssss = rss.RecordCount
If RealOption.Value = True Then
   
   
   If rs.State = 1 Then rs.Close
   rs.Open "select ReelNo,Dates,Weight,GSM,Size,Mill,DateofIssue,[Reel_Issued] from Reel_ReamDetails WHERE ree_ream='Reel' and " & s & " and " & stringyear & "", CON, adOpenKeyset, adLockReadOnly
   For J = 1 To rs.RecordCount
     vs.TextMatrix(J, 0) = J
     vs.TextMatrix(J, 1) = rs!ReelNo
     vs.TextMatrix(J, 2) = rs!Dates
     vs.TextMatrix(J, 3) = rs!weight
     vs.TextMatrix(J, 4) = rs!GSM
     vs.TextMatrix(J, 5) = rs!Size
     vs.TextMatrix(J, 6) = rs!mill
     If rs.Fields("Reel_Issued").Value = True Then
        vs.TextMatrix(J, 7) = rs.Fields("DateofIssue").Value & ""
        vs.TextMatrix(J, 8) = rs.Fields("Reel_Issued").Value & ""
     End If
     
     rec = rec + 1
     If rs.Fields("Reel_Issued").Value = True Then
       Issue = Issue + 1
     End If
     
     rs.MoveNext
   Next
   
   
Else

End If


txtRec = rec
txtIssue = Issue
txtBal = rec - Issue


vs.FormatString = "S.No|ReelNo|Rec.Date|Weight|GSM|Size|Mill|IssueDate|Status"
vs.ColWidth(0) = 800
vs.ColWidth(1) = 1500
vs.ColWidth(2) = 1300
vs.ColWidth(3) = 1300
vs.ColWidth(4) = 1500
vs.ColWidth(5) = 2500
vs.ColWidth(6) = 3000
vs.ColWidth(7) = 1200
vs.ColWidth(8) = 1200


End Sub

Private Sub Form_Load()
    AddItem
    
    rs.Close
    rs.Open "Select * from setup where " & stringyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
    CNSetup
    date1.Text = rs!yarfrom
    date2.Text = rs!yarto
    rs.Close
    
    
    vs.FormatString = "S.No|ReelNo|Weight|GSM|Size|Mill|IssueDate|Status"
    vs.ColWidth(0) = 800
    vs.ColWidth(1) = 1500
    vs.ColWidth(2) = 1300
    vs.ColWidth(3) = 1300
    vs.ColWidth(4) = 1500
    vs.ColWidth(5) = 2000
    vs.ColWidth(6) = 2200
    vs.ColWidth(7) = 1200


    
End Sub
Sub AddItem()
   
   cboGSM.Clear
   cboSize.Clear
   cboMillName.Clear
   
   If rs.State = 1 Then rs.Close
   rs.Open "select Name from Reel_ReamMaster where category='GSM' order by Name", CON, adOpenKeyset, adLockReadOnly
   While rs.EOF = False
        cboGSM.AddItem rs(0)
        rs.MoveNext
   Wend
   cboGSM.AddItem ""
   
   If rs.State = 1 Then rs.Close
   rs.Open "select Name from Reel_ReamMaster where category='Size' order by Name", CON, adOpenKeyset, adLockReadOnly
   While rs.EOF = False
        cboSize.AddItem rs(0)
        rs.MoveNext
   Wend
   
   cboSize.AddItem ""
   
   
   If rs.State = 1 Then rs.Close
   rs.Open "select Name from Reel_ReamMaster where category='Mill' order by Name", CON, adOpenKeyset, adLockReadOnly
   While rs.EOF = False
        cboMillName.AddItem rs(0)
        rs.MoveNext
   Wend
   
   cboMillName.AddItem ""
   
   
   
End Sub

Private Sub vs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
  If vs.Cols = 8 Then
     vs.TextMatrix(vs.RowSel, 6) = Format(Date, "dd/MM/yyyy")
  End If
    
End Sub

