VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSeriesWiseDis 
   Caption         =   "Series Wise Discount"
   ClientHeight    =   9516
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   17604
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   10.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSeriesWiseDis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9516
   ScaleWidth      =   17604
   Begin VB.CheckBox Check1_all 
      BackColor       =   &H8000000E&
      Caption         =   "Click Search For All School"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8136
      TabIndex        =   21
      Top             =   2268
      Width           =   2220
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   840
      Left            =   6984
      Locked          =   -1  'True
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   2772
      Width           =   6828
   End
   Begin VB.ComboBox cboGpSchool 
      Height          =   384
      ItemData        =   "frmSeriesWiseDis.frx":000C
      Left            =   1584
      List            =   "frmSeriesWiseDis.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1404
      Width           =   5304
   End
   Begin VB.TextBox txtSchool 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1584
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2196
      Width           =   5304
   End
   Begin VB.TextBox txtScId 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6948
      TabIndex        =   15
      Top             =   2232
      Width           =   984
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   1584
      MaxLength       =   50
      TabIndex        =   14
      Top             =   108
      Width           =   7284
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   192
      TabIndex        =   13
      Top             =   2748
      Width           =   6585
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   45
         Picture         =   "frmSeriesWiseDis.frx":0010
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   1212
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1284
         Picture         =   "frmSeriesWiseDis.frx":0BF4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2616
         Picture         =   "frmSeriesWiseDis.frx":17D8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
         Width           =   1284
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3924
         Picture         =   "frmSeriesWiseDis.frx":23BC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5280
         Picture         =   "frmSeriesWiseDis.frx":27C9
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   45
         Width           =   1212
      End
   End
   Begin VB.ComboBox cboSeries 
      Height          =   384
      ItemData        =   "frmSeriesWiseDis.frx":33AD
      Left            =   1584
      List            =   "frmSeriesWiseDis.frx":33AF
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   576
      Width           =   5304
   End
   Begin VB.TextBox txtDis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1584
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1008
      Width           =   2244
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5688
      Left            =   72
      TabIndex        =   3
      Top             =   3744
      Width           =   16284
      _cx             =   28723
      _cy             =   10033
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
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
      ForeColorSel    =   4210752
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
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
      ExplorerBar     =   1
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
   Begin MSComCtl2.DTPicker txtDates 
      Height          =   336
      Left            =   6912
      TabIndex        =   19
      Top             =   1404
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   593
      _Version        =   393216
      Format          =   183369729
      CurrentDate     =   39795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Group School :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   144
      TabIndex        =   18
      Top             =   1476
      Width           =   1512
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search School"
      Height          =   372
      Left            =   1584
      TabIndex        =   17
      Top             =   1908
      Width           =   4728
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "School : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   1
      Left            =   180
      TabIndex        =   16
      Top             =   2196
      Width           =   1080
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   144
      Top             =   2700
      Width           =   6720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Series Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   144
      TabIndex        =   12
      Top             =   648
      Width           =   1512
   End
   Begin VB.Label lblName1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   6924
      TabIndex        =   11
      Top             =   624
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   144
      TabIndex        =   10
      Top             =   1056
      Width           =   1260
   End
End
Attribute VB_Name = "frmSeriesWiseDis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean


Private Sub cboGpSchool_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtSchool.SetFocus
End If
  
End Sub

Private Sub cboSeries_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = 13 Then
     txtDis.SetFocus
  End If
  
End Sub
Private Sub cmdAdd_1_Click()


vs.Clear

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False
edit1 = False


fillVs

cboSeries.text = ""
txtDis.text = ""
lblName1.Caption = ""
'txtScId.text = ""
'txtSchool.text = ""
cboGpSchool.text = ""
  

cboSeries.SetFocus

End Sub
Private Sub cmdDelete_3_Click()
Dim lastYrs As String

If (DateDiff("d", Now, SessionLastDate) <= 0) Then
     lastYrs = "current"
Else
     lastYrs = "last"
End If


If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then


If lastYrs = "last" Then
   con_LAST.Execute "delete from  SeriesWiseDiscount where id='" & lblName1.Caption & "'"
Else
   con.Execute "delete from  SeriesWiseDiscount where id='" & lblName1.Caption & "'"
End If

End If

cmdAdd_1_Click

End Sub

Private Sub cmdEdit_4_Click()

edit1 = True

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = True

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdSave_2_Click()

On Error GoTo aa:

Dim lastYrs As String
lastYrs = ""

If (DateDiff("d", Now, SessionLastDate) <= 0) Then
     lastYrs = "current"
Else
     lastYrs = "last"
End If


If cboSeries.text = "" Then
   MsgBox "Select Series Name. ...", vbInformation
   cboSeries.SetFocus
   Exit Sub
End If

Dim scid As String
Dim GpSchool As String

If txtScId.text = "" Then
   scid = "n"
   txtScId.text = "n"
Else
   scid = txtScId.text
End If

If cboGpSchool.text = "" Then
   GpSchool = "n"
   cboGpSchool.text = "n"
Else
   GpSchool = cboGpSchool.text
End If




If edit1 = True Then
   
   If lastYrs = "last" Then
   
     con_LAST.Execute "update [SeriesWiseDiscount] set SeriesName='" & UCase(cboSeries.text) & "',DISCOUNT='" & txtDis.text & "'" & _
    ",party='" & txtParty.text & "',ScName='" & txtSchool.text & "',ScId='" & scid & "',GroupOfSchool='" & GpSchool & "'" & _
    ",UpdatedDate ='" & Format(txtDates.value, "MM/dd/yyyy") & "' where id='" & lblName1.Caption & "'"
    
   
   Else
   
    con.Execute "update [SeriesWiseDiscount] set SeriesName='" & UCase(cboSeries.text) & "',DISCOUNT='" & txtDis.text & "'" & _
    ",party='" & txtParty.text & "',ScName='" & txtSchool.text & "',ScId='" & scid & "',GroupOfSchool='" & GpSchool & "'" & _
    ",UpdatedDate ='" & Format(txtDates.value, "MM/dd/yyyy") & "' where id='" & lblName1.Caption & "'"
   
   End If
   
   Call cmdAdd_1_Click

Else


If RS.State = 1 Then RS.close

Dim str_ As String
str_ = ""

str_ = "substring(Party,1,5)='" & Mid(txtParty.text, 1, 5) & "'"

If cboSeries.text <> "" Then
   str_ = str_ & " and SeriesName='" & UCase(cboSeries.text) & "'"
End If

If cboGpSchool.text <> "" Then
   str_ = str_ & " and GroupOfSchool='" & cboGpSchool.text & "'"
End If

If txtScId.text <> "" Then
   str_ = str_ & " and scid='" & txtScId.text & "'"
End If

If lastYrs = "last" Then
  RS.Open "select * from SeriesWiseDiscount where " & str_, con_LAST
Else
  RS.Open "select * from SeriesWiseDiscount where " & str_, con
End If


If RS.EOF = True Then


If lastYrs = "last" Then

con_LAST.Execute "INSERT INTO  [SeriesWiseDiscount]" & _
           "(SeriesName,DISCOUNT,Party,ScName,ScId,GroupOfSchool,UpdatedDate" & _
     ") Values" & _
           "('" & UCase(cboSeries.text) & "'" & _
           "," & txtDis.text & ",'" & txtParty.text & "','" & txtSchool.text & "','" & scid & "','" & GpSchool & "','" & Format(txtDates.value, "MM/dd/yyyy") & "')"
Else

con.Execute "INSERT INTO  [SeriesWiseDiscount]" & _
           "(SeriesName,DISCOUNT,Party,ScName,ScId,GroupOfSchool,UpdatedDate" & _
     ") Values" & _
           "('" & UCase(cboSeries.text) & "'" & _
           "," & txtDis.text & ",'" & txtParty.text & "','" & txtSchool.text & "','" & scid & "','" & GpSchool & "','" & Format(txtDates.value, "MM/dd/yyyy") & "')"


End If


Call cmdAdd_1_Click
  
Else

   MsgBox "This Series Already Exist. ", vbCritical
   

End If

End If





Screen.MousePointer = vbDefault




fillVs

cboSeries.SetFocus

   
Exit Sub

aa:

Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION

End Sub
Sub fillVs()

Dim lastYrs As String

lastYrs = ""

If (DateDiff("d", Now, SessionLastDate) <= 0) Then
     lastYrs = "current"
Else
     lastYrs = "last"
End If



vs.Cols = 8

k1 = 1
vs.Clear
vs.rows = 1


pcode = Trim(Mid(txtParty.text, 1, 6))

Set rs1 = New ADODB.Recordset


If lastYrs = "last" Then

rs1.Open "select Id,SeriesName,DISCOUNT,LastYrs_Discount,ScName,GroupOfSchool,ScId,UpdatedDate from SeriesWiseDiscountQry_New where substring(Party,1,6)='" & pcode & "' order by SeriesName", con_LAST

Else
rs1.Open "select Id,SeriesName,DISCOUNT,LastYrs_Discount,ScName,GroupOfSchool,ScId,UpdatedDate from SeriesWiseDiscountQry_New where substring(Party,1,6)='" & pcode & "' order by SeriesName", con

End If

While rs1.EOF = False


vs.rows = vs.rows + 1

vs.TextMatrix(k1, 0) = rs1(0)
vs.TextMatrix(k1, 1) = rs1(1)
vs.TextMatrix(k1, 2) = rs1(2)
vs.TextMatrix(k1, 3) = rs1(3) & ""
vs.TextMatrix(k1, 4) = rs1!GroupOfSchool & ""
vs.TextMatrix(k1, 5) = rs1!scname & ""
vs.TextMatrix(k1, 6) = rs1!scid & ""
vs.TextMatrix(k1, 7) = rs1!UpdatedDate & ""


k1 = k1 + 1

rs1.MoveNext

Wend


dis1_ = "Discount "    '& "<= " & SessionLastDate

vs.FormatString = "Id|Series Name|Discount|LstYrs_Dis.|Group Of School|School Name|ScId|UpdatedDt."

vs.TextMatrix(0, 2) = dis1_

vs.ColWidth(0) = 800
vs.ColWidth(1) = 2400
vs.ColWidth(2) = 1800
vs.ColWidth(3) = 1400
vs.ColWidth(4) = 3300
vs.ColWidth(5) = 3300
vs.ColWidth(6) = 900


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
   Unload Me
End If

End Sub

Private Sub Form_Load()

txtDates.value = Format(Date, "dd/Mm/yyyy")

Me.Left = 1000
Me.top = 1000
Me.Width = 15280
Me.Height = 9980

txtParty.text = PopUpValue6

PopUpValue6 = ""



If RS.State = 1 Then RS.close
RS.Open "select distinct SerName from books  order by SerName", con
While RS.EOF = False

If Not IsNull(RS(0)) Then
   cboSeries.AddItem Trim(RS(0))
End If

RS.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "SELECT GroupOfSchool FROM collegeView where len(GroupOfSchool)>0 group by GroupOfSchool order by GroupOfSchool", CON_blue
While RS.EOF = False

If Not IsNull(RS(0)) Then
   cboGpSchool.AddItem Trim(RS(0))
End If

RS.MoveNext
Wend


txtrem.text = ""

Set rs1 = New ADODB.Recordset
rs1.Open "select PartyRemarks from sledger where subledger='" & txtParty.text & "'", con
If rs1.EOF = False Then
   txtrem.text = rs1.Fields("PartyRemarks").value & ""
End If


fillVs

'cboSeries.SetFocus

End Sub
Private Sub txtDis_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   cboGpSchool.SetFocus
End If
  
End Sub
Private Sub txtschool_GotFocus()

If RS.State = 1 Then RS.close
If PopUpValue1 <> "" Then
   
    If (Check1_all.value = 1) Then
    
     txtScId.text = PopUpValue1
     txtSchool.text = PopUpValue2 & ", " & PopUpValue3
    
    Else
    
     txtScId.text = PopUpValue2
     txtSchool.text = PopUpValue1
   
    End If
    
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""

End If

End Sub

Private Sub txtschool_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   
 If (Check1_all.value = 1) Then
   Screen.MousePointer = vbHourglass
   tblNo = 9
   frmSearchItem.Show 1
   Screen.MousePointer = vbDefault
 Else
 
    value = "select ScName,Scid from ORDERA where partyname='" & txtParty.text & "' group by ScName,Scid order by ScName"
    popuplist10 value, con
 
 End If
 
End If

If KeyCode = 13 Then
   cmdSave_2_Click
End If
  

End Sub

Private Sub vs_DblClick()

lblName1.Caption = vs.TextMatrix(vs.RowSel, 0)
cboSeries.text = vs.TextMatrix(vs.RowSel, 1)
txtDis.text = vs.TextMatrix(vs.RowSel, 2)

cboGpSchool.text = vs.TextMatrix(vs.RowSel, 4)

txtSchool.text = vs.TextMatrix(vs.RowSel, 5)
txtScId.text = vs.TextMatrix(vs.RowSel, 6)


If LCase(UserName) = "admin" Then

cmdEdit_4.Enabled = True
cmdDelete_3.Enabled = False
cmdSave_2.Enabled = False

End If




End Sub
