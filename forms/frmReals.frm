VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmReals 
   Caption         =   "Reel/Ream"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   17295
   WindowState     =   2  'Maximized
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   1350
      TabIndex        =   21
      Top             =   7275
      Width           =   8835
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   45
         Picture         =   "frmReals.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1260
         Picture         =   "frmReals.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2520
         Picture         =   "frmReals.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3750
         Picture         =   "frmReals.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7500
         Picture         =   "frmReals.frx":27B9
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6255
         Picture         =   "frmReals.frx":339D
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4980
         Picture         =   "frmReals.frx":3F81
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.TextBox txtWeight 
      Height          =   360
      Left            =   9150
      TabIndex        =   7
      Top             =   2475
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txtRealNo 
      Height          =   360
      Left            =   9150
      TabIndex        =   6
      Top             =   2025
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   5325
      TabIndex        =   18
      Top             =   2025
      Visible         =   0   'False
      Width           =   2790
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
         Height          =   465
         Left            =   1275
         TabIndex        =   5
         Top             =   225
         Width           =   1320
      End
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
         Height          =   390
         Left            =   150
         TabIndex        =   4
         Top             =   225
         Value           =   -1  'True
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdmill 
      Caption         =   "Add"
      Height          =   390
      Left            =   4350
      TabIndex        =   16
      Top             =   2475
      Width           =   840
   End
   Begin VB.ComboBox cboMillName 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2550
      Width           =   2865
   End
   Begin VB.CommandButton cmdAdd2 
      Caption         =   "Add"
      Height          =   390
      Left            =   4350
      TabIndex        =   14
      Top             =   2025
      Width           =   840
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2100
      Width           =   2865
   End
   Begin VB.CommandButton cmdAdd1 
      Caption         =   "Add"
      Height          =   390
      Left            =   4350
      TabIndex        =   12
      Top             =   1575
      Width           =   840
   End
   Begin VB.ComboBox cboGSM 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1650
      Width           =   2865
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   315
      Left            =   1425
      TabIndex        =   0
      Top             =   1200
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
      Height          =   4050
      Left            =   1425
      TabIndex        =   29
      Top             =   3000
      Width           =   5835
      _cx             =   10292
      _cy             =   7144
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
      ForeColorFixed  =   -2147483630
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
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmReals.frx":4B65
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
   Begin VB.Label lblHead 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   1425
      TabIndex        =   28
      Top             =   75
      Width           =   7890
   End
   Begin VB.Label Label7 
      Caption         =   "Date"
      Height          =   240
      Left            =   375
      TabIndex        =   20
      Top             =   1275
      Width           =   840
   End
   Begin VB.Label Label6 
      Caption         =   "Weight :"
      Height          =   240
      Left            =   8250
      TabIndex        =   19
      Top             =   2550
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lblReel 
      Caption         =   "Real No :"
      Height          =   240
      Left            =   8250
      TabIndex        =   17
      Top             =   2100
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label4 
      Caption         =   "Mill Name  :"
      Height          =   240
      Left            =   375
      TabIndex        =   15
      Top             =   2550
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Size :"
      Height          =   240
      Left            =   375
      TabIndex        =   13
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "GSM :"
      Height          =   240
      Left            =   375
      TabIndex        =   11
      Top             =   1725
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Entry No :"
      Height          =   240
      Left            =   375
      TabIndex        =   10
      Top             =   750
      Width           =   840
   End
   Begin VB.Label lblno 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1425
      TabIndex        =   9
      Top             =   750
      Width           =   1440
   End
End
Attribute VB_Name = "frmReals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dataEdit As Boolean

Private Sub cmdAdd_1_Click()
'lblno = MaxSNo("Reel_ReamDetails", "EntryNo")

MaxRecNo

dataEdit = False

cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False

txtDate.SetFocus
txtDate.Text = Format(txtDate, "dd/MM/yyyy")
cboGSM.ListIndex = -1
cboSize.ListIndex = -1
cboMillName.ListIndex = -1
txtRealNo = ""
txtWeight = 0

Setgrid


End Sub

Private Sub cmdAdd1_Click()
Paper_Master = "GSM"
frmReelMasters.Show 1
End Sub

Private Sub cmdAdd2_Click()
Paper_Master = "Size"
frmReelMasters.Show 1

End Sub

Private Sub cmdDelete_3_Click()
If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
      CON.Execute "delete from Reel_ReamDetails where [EntryNo]=" & Val(lblno.Caption) & " and " & stringyear
      Call cmdAdd_1_Click
End If
End Sub

Private Sub cmdEdit_4_Click()
dataEdit = True
cmdDelete_3.Enabled = True
cmdSave_2.Enabled = True

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdmill_Click()
Paper_Master = "Mill"
frmReelMasters.Show 1


End Sub
Sub AddItem()
   
   cboGSM.Clear
   cboSize.Clear
   cboMillName.Clear
   
   If rs.State = 1 Then rs.Close
   rs.Open "select Name from Reel_ReamMaster where " & stringyear & " and  category='GSM' order by Name", CON, adOpenKeyset, adLockReadOnly
   While rs.EOF = False
        cboGSM.AddItem rs(0)
        rs.MoveNext
   Wend
   
   
   If rs.State = 1 Then rs.Close
   rs.Open "select Name from Reel_ReamMaster where " & stringyear & " and  category='Size' order by Name", CON, adOpenKeyset, adLockReadOnly
   While rs.EOF = False
        cboSize.AddItem rs(0)
        rs.MoveNext
   Wend
   
   
   If rs.State = 1 Then rs.Close
   rs.Open "select Name from Reel_ReamMaster where " & stringyear & " and  category='Mill' order by Name", CON, adOpenKeyset, adLockReadOnly
   While rs.EOF = False
        cboMillName.AddItem rs(0)
        rs.MoveNext
   Wend
   
   
   
   
End Sub

Private Sub cmdPrint_Click()
frmReel_ReamRpt.Show
Unload frmReals
End Sub

Private Sub cmdSave_2_Click()

Dim reel As String

If RealOption.Value = True Then
reel = "Reel"
Else
reel = "Ream"
End If



vs.Cols = 3

If rs.State = 1 Then rs.Close
rs.Open "select * from Reel_ReamDetails where " & stringyear & " and  [EntryNo]=" & Val(lblno.Caption) & "", CON, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then
   
   CON.Execute "delete from Reel_ReamDetails where " & stringyear & " and  [EntryNo]=" & Val(lblno.Caption) & " and issue_receive='Receive' and " & stringyear
   If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
      
   For d1 = 1 To vs.Rows - 1
   
   If vs.TextMatrix(d1, 1) <> "" Then
      
      CON.Execute "insert into Reel_ReamDetails([EntryNo],[Dates],[GSM],[Size],[Mill],[Ree_Ream],[ReelNo],[Weight],[fyear],[Createdby],setupid,Issue_Receive) " & _
      " values(" & Val(lblno.Caption) & ",'" & Format(txtDate, "MM/dd/yyyy") & "'," & _
      "'" & cboGSM & "','" & cboSize & "','" & cboMillName & "','" & reel & "','" & vs.TextMatrix(d1, 1) & "'," & Val(vs.TextMatrix(d1, 2)) & ",'" & main.session & "'," & _
      "'" & main.username & "','" & main.setupid & "','" & receie_issue & "')"
   End If
   
   Next
    
    End If
    
Else

   MaxRecNo

   If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
      
    
    
   For d1 = 1 To vs.Rows - 1
   
   If vs.TextMatrix(d1, 1) <> "" Then
      
      CON.Execute "insert into Reel_ReamDetails([EntryNo],[Dates],[GSM],[Size],[Mill],[Ree_Ream],[ReelNo],[Weight],[fyear],[Createdby],setupid,Issue_Receive) " & _
      " values(" & Val(lblno.Caption) & ",'" & Format(txtDate, "MM/dd/yyyy") & "'," & _
      "'" & cboGSM & "','" & cboSize & "','" & cboMillName & "','" & reel & "','" & vs.TextMatrix(d1, 1) & "'," & Val(vs.TextMatrix(d1, 2)) & ",'" & main.session & "'," & _
      "'" & main.username & "','" & main.setupid & "','" & receie_issue & "')"
   End If
   
   Next

    
  End If

   
End If
   
   Call cmdAdd_1_Click
'End If

End Sub
Sub MaxRecNo()

If rs.State = 1 Then rs.Close
rs.Open "select max(entryno) from Reel_ReamDetails where " & stringyear & " and  issue_receive='" & receie_issue & "'", CON, adOpenKeyset, adLockReadOnly
If IsNull(rs(0)) Then
   lblno.Caption = 1
Else
  lblno.Caption = rs(0) + 1
End If

End Sub

Private Sub cmdSearch_Click()
popuplist10 "select  distinct EntryNo,Dates,[GSM],[Size],[Mill] as [Mill Name] FROM [Reel_ReamDetails] where " & stringyear & " and  issue_Receive='" & receie_issue & "' and " & stringyear & " order by EntryNo", CON
End Sub
Private Sub cmdSearch_GotFocus()
   
   If PopUpValue1 <> "" Then
      searchData
      PopUpValue1 = ""
   End If

End Sub
Private Sub Form_Activate()
AddItem
End Sub
Sub searchData()

      
    If rs.State = 1 Then rs.Close
    rs.Open "select [EntryNo],[Dates],[GSM],[Size],[Mill],[Ree_Ream],[ReelNo],[weight] from Reel_ReamDetails where " & stringyear & " and  issue_Receive='" & receie_issue & "' and [entryNo]=" & PopUpValue1 & "", CON, adOpenKeyset, adLockReadOnly
    If rs.EOF = False Then
       cmdDelete_3.Enabled = False
       cmdSave_2.Enabled = False
       
       lblno.Caption = PopUpValue1
       txtDate.Text = rs!Dates
       cboGSM.Text = rs!GSM
       cboSize.Text = rs!Size
       cboMillName.Text = rs!mill
       If rs!Ree_Ream = "Reel" Then
          RealOption.Value = True
       Else
          ReamOption.Value = True
       End If
       
       Setgrid
       
       'txtRealNo = rs![ReelNo]
       'txtWeight = rs![weight]
       
       vs.Rows = 1
       
       For i = 1 To rs.RecordCount
          vs.Rows = vs.Rows + 1
          vs.TextMatrix(i, 0) = i
          vs.TextMatrix(i, 1) = rs.Fields("ReelNo").Value
          vs.TextMatrix(i, 2) = rs.Fields("weight").Value
          rs.MoveNext
       Next
       
       
       txtDate.SetFocus
       
    End If


End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
     ' SendKeys "{tab}"
    If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("vs")) Then
     SendKeys ("{TAB}")
    End If
   End If
   
   
   
End Sub

Private Sub Form_Load()
'lblno = MaxSNo("Reel_ReamDetails", "EntryNo")

MaxRecNo

dataEdit = False

txtDate = Format(Date, "mm/DD/yyyy")

If receie_issue = "Receive" Then

lblHead.Caption = "Receive (Reel)"
'ReamOption.Value = False
'txtRealNo.Visible = True
'lblReel.Visible = True

Else

'ReamOption.Value = True
'txtRealNo.Visible = False
'lblReel.Visible = True
lblHead.Caption = "Issue (Reel)"

End If

Setgrid



End Sub
Sub Setgrid()
    vs.Clear
    
    vs.Cols = 3
    vs.FormatString = "S.N.|Reel No|Weight"
    
    vs.ColWidth(0) = 800
    vs.ColWidth(1) = 1500
    vs.ColWidth(2) = 1200
    
End Sub
Private Sub RealOption_Click()
If RealOption.Value = True Then
   txtRealNo.Enabled = True
Else
   txtRealNo.Text = ""
   txtRealNo.Enabled = False
   
End If
End Sub

Private Sub ReamOption_Click()
If RealOption.Value = True Then
   txtRealNo.Enabled = True
Else
   txtRealNo.Enabled = False
End If
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       
       If vs.Col = 1 Then
         
       If rs.State = 1 Then rs.Close
       rs.Open "select * from Reel_ReamDetails where " & stringyear & " and  ReelNo='" & vs.TextMatrix(vs.RowSel, 1) & "'", CON, adOpenKeyset, adLockReadOnly
       
       If rs.EOF = False Then
          MsgBox "Reen No Already Exist ...", vbInformation
          SendKeys "{right}"
       Else
          vs.TextMatrix(vs.RowSel, 0) = vs.Row
          SendKeys "{right}"
       End If
       
       Else
         SendKeys "{home}"
         SendKeys "{down}"
         
       End If
    End If
End Sub

