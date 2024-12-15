VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSchoolPay 
   ClientHeight    =   8436
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8436
   ScaleWidth      =   13140
   Begin VB.Frame panel 
      Caption         =   "Donation Entry"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7620
      Left            =   315
      TabIndex        =   0
      Top             =   180
      Width           =   12750
      Begin VB.ComboBox cboagent 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1215
         Width           =   4020
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8385
         TabIndex        =   17
         Top             =   5955
         Width           =   1275
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   495
         TabIndex        =   9
         Top             =   6615
         Width           =   8985
         Begin VB.CommandButton cmdExit_12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            Height          =   660
            Left            =   7755
            Picture         =   "frmSchoolPay.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   75
            Width           =   1125
         End
         Begin VB.CommandButton cmdPrint_7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print (Agent Wise)"
            Height          =   660
            Left            =   5880
            Picture         =   "frmSchoolPay.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   90
            Width           =   1800
         End
         Begin VB.CommandButton cmdEdit_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   660
            Left            =   3450
            Picture         =   "frmSchoolPay.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   75
            Width           =   1125
         End
         Begin VB.CommandButton cmdDelete_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   660
            Left            =   2265
            Picture         =   "frmSchoolPay.frx":1C0A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   75
            Width           =   1125
         End
         Begin VB.CommandButton cmdSave_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   660
            Left            =   1140
            Picture         =   "frmSchoolPay.frx":27EE
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   75
            Width           =   1065
         End
         Begin VB.CommandButton cmdAdd_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   660
            Left            =   75
            Picture         =   "frmSchoolPay.frx":33D2
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   75
            Width           =   1005
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   660
            Left            =   4635
            Picture         =   "frmSchoolPay.frx":3FB6
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   90
            Width           =   1170
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Output Weight"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   255
         TabIndex        =   4
         Top             =   8985
         Visible         =   0   'False
         Width           =   465
         Begin VB.TextBox txtRawAndCasting 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3690
            TabIndex        =   5
            Text            =   "0"
            Top             =   1035
            Width           =   1320
         End
         Begin VB.Shape Shape2 
            Height          =   585
            Left            =   75
            Top             =   1365
            Width           =   3150
         End
         Begin VB.Label Label9 
            Caption         =   "Receiving Semi Finish &&  Finish"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   150
            TabIndex        =   8
            Top             =   1515
            Width           =   3060
         End
         Begin VB.Shape Shape1 
            Height          =   615
            Left            =   0
            Top             =   1590
            Width           =   3135
         End
         Begin VB.Label Label1 
            Caption         =   "Receiving  from casting"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   150
            TabIndex        =   7
            Top             =   885
            Width           =   2325
         End
         Begin VB.Label Label1 
            Caption         =   "Raw Issue Weight"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   135
            TabIndex        =   6
            Top             =   570
            Width           =   1635
         End
      End
      Begin VB.TextBox txtEntNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1605
         TabIndex        =   3
         Top             =   600
         Width           =   1380
      End
      Begin VB.ComboBox cboCollege 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         Top             =   1635
         Width           =   9735
      End
      Begin VB.CheckBox Check1_same 
         BackColor       =   &H80000013&
         Caption         =   "Same S.N.        (Press F1 For Same S. No)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   2130
         Width           =   5595
      End
      Begin Crystal.CrystalReport CR 
         Left            =   11790
         Top             =   7110
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker Dates 
         Height          =   315
         Left            =   1605
         TabIndex        =   19
         Top             =   915
         Width           =   1395
         _ExtentX        =   2455
         _ExtentY        =   572
         _Version        =   393216
         Format          =   156368897
         CurrentDate     =   39500
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   45
         Top             =   9180
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   593
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "filedsn=saru;"
         OLEDBString     =   "filedsn=saru;"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "ItemMaster"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   3465
         Left            =   90
         TabIndex        =   20
         Top             =   2415
         Width           =   12600
         _cx             =   22225
         _cy             =   6112
         _ConvInfo       =   1
         Appearance      =   2
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   7917545
         ForeColorFixed  =   -2147483630
         BackColorSel    =   13432946
         ForeColorSel    =   16711680
         BackColorBkg    =   -2147483626
         BackColorAlternate=   16777215
         GridColor       =   8388608
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   800
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSchoolPay.frx":4B9A
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
         OutlineCol      =   2
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
         Begin VB.Frame VsFrame 
            Height          =   2370
            Left            =   60
            TabIndex        =   21
            Top             =   1740
            Visible         =   0   'False
            Width           =   4155
            Begin MSDataListLib.DataCombo cboItem 
               Height          =   2310
               Left            =   60
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   60
               Width           =   4125
               _ExtentX        =   7281
               _ExtentY        =   3958
               _Version        =   393216
               Appearance      =   0
               Style           =   1
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   915
         Left            =   450
         Top             =   6570
         Width           =   9105
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   90
         TabIndex        =   28
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 For Search "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         TabIndex        =   27
         Top             =   360
         Width           =   2505
      End
      Begin VB.Label Label1 
         Caption         =   "Total :"
         Height          =   195
         Index           =   6
         Left            =   7365
         TabIndex        =   26
         Top             =   6015
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   25
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ent. No :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   24
         Top             =   615
         Width           =   1530
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         Caption         =   "School Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   23
         Top             =   1635
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmSchoolPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Dim RS As New ADODB.Recordset
Dim b1 As Boolean
Private Sub cboagent_Click()

cboCollege.Clear
If rs1.State = 1 Then rs1.close

rs1.Open "SELECT College.College, College.city, College.district " & _
" FROM College INNER JOIN DISTRICTS ON College.district = DISTRICTS.DISTRICTNAME where College.fyear='" & session & "' and College.setupid=" & setupid & " and DISTRICTS.AGENTNAME='" & cboagent.Text & "' order by College.College", con, adOpenForwardOnly, adLockReadOnly
While rs1.EOF = False
   If rs1.Fields(1).value = rs1.Fields(2).value Then
     cboCollege.AddItem rs1.Fields(0).value & " " & rs1.Fields(1).value
   Else
     cboCollege.AddItem rs1.Fields(0).value & " " & rs1.Fields(1).value & " " & rs1.Fields(2).value
   End If
   rs1.MoveNext
Wend


End Sub

Private Sub cboagent_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboCollege.SetFocus
End Sub

Private Sub cboCollege_GotFocus()
SendKeys "{f4}"
End Sub

Private Sub cboCollege_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then vs.SetFocus
End Sub

Private Sub cmdAdd_1_Click()
MaxNo

cboagent.ListIndex = -1
cboCollege = ""

txtTotal = ""

vs.Clear
setWidth


cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True

Check1_same.value = 0
txtEntNo.SetFocus



End Sub

Private Sub cmdDelete_3_Click()
If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
   
   con.Execute "delete from DonationA where " & stringyear & " and EntNo=" & txtEntNo & ""
   con.Execute "delete from DonationB where " & stringyear & " and EntNo=" & txtEntNo & ""
   
   cmdAdd_1_Click
End If

End Sub

Private Sub cmdEdit_4_Click()

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = True
cmdSave_2.SetFocus
  
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub


Sub Total()

txtTotal.Text = 0

For J = 1 To vs.rows - 1
If vs.TextMatrix(J, 4) <> "" Then
 
 txtTotal.Text = (Val(txtTotal.Text) + Val(vs.TextMatrix(J, 4)))
 
End If
Next

End Sub


Sub MaxNo()

If rs1.State = 1 Then rs1.close
rs1.Open "select max(EntNo) from DonationA", con, adOpenDynamic, adLockReadOnly
If IsNull(rs1(0)) Then
   txtEntNo = 1
 Else
   txtEntNo = Val(rs1(0)) + 1
End If
   

End Sub
Private Sub cmdPrint_7_Click()

DSNNew

Screen.MousePointer = vbHourglass

cr.Reset
cr.ReportFileName = rptPath & "/Donation.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.ReplaceSelectionFormula "{donationA.AgentName}='" & cboagent.Text & "'"
cr.WindowShowPrintSetupBtn = True
cr.WindowState = crptMaximized
cr.Action = 1





Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSave_2_Click()

If RS.State = 1 Then RS.close
RS.Open "select * from DonationA where " & stringyear & " and EntNo=" & txtEntNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
   RS.AddNew
Else
  con.Execute "delete from Donationb where " & stringyear & " and EntNo=" & txtEntNo & ""
End If

RS!entno = txtEntNo
RS!dates = dates.value
RS!agentname = Trim(cboagent)
RS!SchoolName = Trim(cboCollege)
RS!fyear = session
RS!setupid = setupid

RS.update


If RS.State = 1 Then RS.close
RS.Open "select * from DonationB ", con, adOpenDynamic, adLockOptimistic
For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 3) <> "" Then

RS.AddNew
RS!entno = txtEntNo
RS!Bookcode = IIf(vs.TextMatrix(I, 1) = "", "-", vs.TextMatrix(I, 1))
RS!Teacher = vs.TextMatrix(I, 3)
RS!amount = Val(vs.TextMatrix(I, 4))
RS!Others = vs.TextMatrix(I, 5)
RS!sno = vs.TextMatrix(I, 0)
RS!Status = vs.TextMatrix(I, 6)

RS!fyear = session
RS!setupid = setupid


RS.update

End If

Next

MsgBox "Data Saved ...", vbInformation

cmdSave_2.Enabled = False

End Sub
Sub searchData()

'On Error Resume Next

vs.Clear
setWidth



If RS.State = 1 Then RS.close
RS.Open "select * from DonationA where " & stringyear & " and EntNo=" & txtEntNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   
    cmdSave_2.Enabled = False
    cmdEdit_4.Enabled = True
    cmdDelete_3.Enabled = False
   
    txtEntNo = RS!entno
    dates.value = RS!dates
    cboagent = RS!agentname
    cboCollege = RS!SchoolName

End If


If RS.State = 1 Then RS.close
RS.Open "SELECT db.EntNo,db.BookCode,db.Teacher,db.Amount,b.BOOKNAME,db.Others,db.sno,db.Status FROM DonationB db INNER JOIN BOOKS b ON db.BookCode = b.BOOKCODE where db.fyear='" & session & "' and db.setupid=" & setupid & " and db.EntNo=" & txtEntNo & "", con, adOpenDynamic, adLockOptimistic
For I = 1 To RS.RecordCount


vs.TextMatrix(I, 0) = I
vs.TextMatrix(I, 1) = RS!Bookcode
vs.TextMatrix(I, 2) = RS!Bookname

vs.TextMatrix(I, 3) = RS!Teacher
vs.TextMatrix(I, 4) = RS!amount
vs.TextMatrix(I, 5) = RS!Others
vs.TextMatrix(I, 0) = RS!sno
vs.TextMatrix(I, 6) = RS!Status

RS.MoveNext


Next



vs.MergeCells = flexMergeFixedOnly
vs.MergeCol(0) = True

Total

End Sub


Private Sub Command1_Click()

Screen.MousePointer = vbHourglass
DSNNew

cr.Reset
cr.ReportFileName = rptPath & "/Donation_slip.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.ReplaceSelectionFormula "{donationA.EntNo}=" & txtEntNo & ""
cr.WindowShowPrintSetupBtn = True
cr.WindowState = crptMaximized
cr.Action = 1

Screen.MousePointer = vbDefault


End Sub

Private Sub dates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboagent.SetFocus
End Sub

Private Sub Form_Activate()
BackColorFrom Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 112 Then
   
 
 If Check1_same.value = 1 Then
    Check1_same.value = 0
 Else
    Check1_same.value = 1
 End If

End If



End Sub

Private Sub Form_Load()
Me.Top = 100
Me.Left = 100

Me.Width = 12800
Me.Height = 8500

setWidth
MaxNo

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT agentname" & _
" FROM AGENTMASTER where " & stringyear & " order by agentname", con, adOpenForwardOnly, adLockReadOnly
While rs1.EOF = False

cboagent.AddItem rs1(0)

rs1.MoveNext
Wend
 
 
Screen.MousePointer = Default
dates.value = Date
 


End Sub

Sub setWidth()
 vs.Cols = 7
 vs.FormatString = "S.No|Book Code|Book Name|Teacher Name|>Amount|Others|Status"
 vs.ColWidth(0) = 700
 vs.ColWidth(1) = 1050
 vs.ColWidth(2) = 3800
 vs.ColWidth(3) = 2400
 vs.ColWidth(4) = 1200
 vs.ColWidth(5) = 1600
 vs.ColWidth(6) = 800
 
End Sub
Private Sub Form_Resize()

panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub txtEntNo_GotFocus()
If PopUpValue1 <> "" Then
   txtEntNo = PopUpValue1
   searchData
   PopUpValue1 = ""
End If
End Sub

Private Sub txtEntNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    popuplistModel10 "Select EntNo,Dates,AgentName from DonationA order by EntNo ", con
End If
End Sub

Private Sub txtEntNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dates.SetFocus
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    popuplistModel10 "Select EntNo,Dates,AgentName from DonationA order by EntNo ", con
End If
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)



If KeyCode = 13 Then

 If vs.Col = 1 Then

    If RS.State = 1 Then RS.close
    RS.Open "select * from BOOKS where " & stringyear & " and BOOKCODE='" & vs.TextMatrix(vs.RowSel, 1) & "'", con
    If RS.EOF = False Then
       
      
            
    '----------------------------------------------
      
    If vs.Row > 1 Then
      
      If Check1_same.value = 0 Then
          vs.TextMatrix(vs.RowSel, 0) = (vs.TextMatrix(vs.RowSel - 1, 0) + 1)
        Else
          
          If Check1_same.value = 1 Then
          
          If b1 = False Then
             vs.TextMatrix(vs.RowSel, 0) = vs.Row
             b1 = True
          End If
        
          If b1 = True Then
             vs.TextMatrix(vs.RowSel, 0) = vs.TextMatrix(vs.RowSel - 1, 0)
          End If
        
       End If
         
    End If
       
   Else
        vs.TextMatrix(vs.RowSel, 0) = vs.Row
        
   End If
       
       vs.TextMatrix(vs.RowSel, 1) = UCase(vs.TextMatrix(vs.RowSel, 1))
       vs.TextMatrix(vs.RowSel, 2) = UCase(RS.Fields("bookname").value)
       
       SendKeys "{right}"
       SendKeys "{right}"
       
       vs.Editable = flexEDKbdMouse
       End If

ElseIf vs.Col = 3 Then
       SendKeys "{right}"
ElseIf vs.Col = 4 Then
       SendKeys "{right}"
       Total
ElseIf vs.Col = 5 Then
       vs.TextMatrix(vs.RowSel, 5) = UCase(vs.TextMatrix(vs.RowSel, 5))
       SendKeys "{right}"
       Total
       vs.Editable = flexEDKbdMouse

ElseIf vs.Col = 6 Then
       
       SendKeys "{home}"
       SendKeys "{down}"
       
End If



End If


''
End Sub
''
Private Sub vs_Click()

End Sub

Private Sub vs_SelChange()
If vs.Col = 3 Then
'   vs.Editable = flexEDNone
Else
'   vs.Editable = flexEDKbdMouse
End If
vs.TextMatrix(vs.RowSel, 0) = vs.Row
End Sub


