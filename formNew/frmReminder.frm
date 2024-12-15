VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReminder 
   Caption         =   "Reminder Message"
   ClientHeight    =   8388
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   13152
   Icon            =   "frmReminder.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8388
   ScaleWidth      =   13152
   Begin VB.CheckBox Check1_All 
      Caption         =   "Show All Reminder Message..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2430
      TabIndex        =   13
      Top             =   90
      Width           =   2490
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Reminder..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   13065
      Begin VB.TextBox txtReminderClseRem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   1890
         Width           =   12075
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   465
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2340
         Width           =   960
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Default         =   -1  'True
         Height          =   465
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2340
         Width           =   960
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   1260
         Width           =   12075
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   375
         Left            =   135
         TabIndex        =   0
         Top             =   585
         Width           =   1455
         _ExtentX        =   2582
         _ExtentY        =   656
         _Version        =   393216
         Format          =   500760577
         CurrentDate     =   43717
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Reminder Close Remarks"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   135
         TabIndex        =   14
         Top             =   1665
         Width           =   2220
      End
      Begin VB.Label lblId 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   1665
         TabIndex        =   11
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Reminder Message"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   135
         TabIndex        =   9
         Top             =   1035
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   135
         TabIndex        =   8
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.CheckBox Check1_Add 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Reminder..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Width           =   2220
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7680
      Left            =   0
      TabIndex        =   4
      Top             =   630
      Width           =   13080
      _cx             =   23072
      _cy             =   13547
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      ForeColorSel    =   16711680
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
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   450
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmReminder.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      WordWrap        =   -1  'True
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click on Message For Edit ..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5175
      TabIndex        =   10
      Top             =   135
      Width           =   3540
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key To Close Message .."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8910
      TabIndex        =   5
      Top             =   135
      Visible         =   0   'False
      Width           =   3360
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub fill()

''On Error Resume Next

 On Error GoTo aaa:


 Dim rsfill As ADODB.Recordset
 Set rsfill = New ADODB.Recordset
 
 Screen.MousePointer = vbHourglass
 If Check1_all.value = 1 Then
    rsfill.Open "select Id,CONVERT(varchar(50),INVOICEDATE,103) as INVOICEDATE,Remarks,CloseRemarks,Show,UserName from reminderTbl order by Id,INVOICEDATE", con
 Else
    'rsfill.Open "select Id,CONVERT(varchar(50),INVOICEDATE,103) as INVOICEDATE,Remarks,CloseRemarks,Show,UserName from reminderTbl where (Show='n' and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" & Date & "',103)) order by Id,INVOICEDATE", con
    rsfill.Open "select Id,CONVERT(varchar(50),INVOICEDATE,103) as INVOICEDATE,Remarks,CloseRemarks,Show,UserName from reminderTbl where (Show='n' and convert(smalldatetime,invoicedate,103) <= convert(smalldatetime,'" & Date & "',103)) order by Id,INVOICEDATE", con
 End If
 Set vs.DataSource = rsfill
 
 '======================================
 For I = 0 To vs.rows - 1
 
 If K = 0 Then
    vs.Cell(flexcpBackColor, I, 0) = vbWhite
    vs.Cell(flexcpBackColor, I, 1) = vbWhite
    vs.Cell(flexcpBackColor, I, 2) = vbWhite
    vs.Cell(flexcpBackColor, I, 3) = vbWhite
    vs.Cell(flexcpBackColor, I, 4) = vbWhite
    vs.Cell(flexcpBackColor, I, 5) = vbWhite
    
    K = 1
   Else
    vs.Cell(flexcpBackColor, I, 0) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 1) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 2) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 3) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 4) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 5) = &HC0FFFF
    
    K = 0
    
  End If
 
 Next
 
 
 
 
 
 
 '======================================
 
 vs.FormatString = "Id|Date|Reminder Message|Reminder Close Message|Show|UserName"
 
 vs.ColWidth(0) = 700
 vs.ColWidth(1) = 1200
 vs.ColWidth(2) = 4500
 vs.ColWidth(3) = 4500
 
 Screen.MousePointer = vbDefault
 
 
Exit Sub

aaa:

MsgBox "" & err.DESCRIPTION

End Sub

Private Sub Check1_Add_Click()

If Check1_Add.value = 1 Then
   Frame1.Visible = True
   txtDate.SetFocus
   lblId.Caption = ""
   txtrem.text = ""
   txtReminderClseRem.text = ""
   cmdSave.Caption = "Add"
Else
   Frame1.Visible = False
End If

End Sub

Private Sub Check1_all_Click()

If Check1_all.value = 1 Then
   Frame1.Visible = False
   fill
Else
   fill
End If

End Sub

Private Sub cmdExit_Click()

txtrem.text = ""
txtReminderClseRem.text = ""

Frame1.Visible = False

End Sub
Private Sub cmdSave_Click()

On Error GoTo aa:

If txtrem.text <> "" Then
   
If lblId.Caption = "" Then

   con.Execute "insert into reminderTbl(invoicedate,remarks,UserName,CloseRemarks) values('" & Format(txtDate.value, "MM/dd/yyyy") & "','" & txtrem.text & "','" & UserName & "','" & Trim(txtReminderClseRem.text) & "')"
   txtrem.text = ""
   fill
Else
   If lblId <> "" Then
     con.Execute "update reminderTbl set invoiceDate='" & Format(txtDate.value, "MM/dd/yyyy") & "',username='" & UserName & "',remarks='" & txtrem & "' where id=" & lblId & ""
     If txtReminderClseRem <> "" Then
        con.Execute "update reminderTbl set CloseRemarks='" & Trim(txtReminderClseRem) & "',show='y' where id=" & lblId & ""
     End If
     fill
   End If

End If
   
Else
   MsgBox "Enter Message...", vbInformation
End If

Exit Sub

aa:

MsgBox "" & err.DESCRIPTION


End Sub
Private Sub Form_Load()
Me.top = 10
Me.Left = 10
Me.Height = 9015
Me.Width = 13275

txtDate.value = Format(Date, "dd/MM/yyyy")

If Right(session, 2) >= 20 Then
fill
End If

End Sub
Private Sub vs_DblClick()
   Frame1.Visible = True
   lblId.Caption = vs.TextMatrix(vs.RowSel, 0)
   txtDate.value = vs.TextMatrix(vs.RowSel, 1)
   txtrem.text = vs.TextMatrix(vs.RowSel, 2)
   cmdSave.Caption = "Update"
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      If MsgBox("Want to Remove Message ?", vbQuestion + vbYesNo) = vbYes Then
         con.Execute "update reminderTbl set show='y' where id=" & vs.TextMatrix(vs.RowSel, 0) & ""
         fill
      End If
   End If
End Sub
