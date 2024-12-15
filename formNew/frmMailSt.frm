VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMailSt 
   Caption         =   "Mail Status"
   ClientHeight    =   10224
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   20472
   Icon            =   "frmMailSt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10224
   ScaleWidth      =   20472
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUpdateCC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit CC Mail"
      Height          =   540
      Left            =   7785
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   90
      Width           =   1365
   End
   Begin VB.CommandButton cmdBlukMail 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Bluk Mail Edit"
      Height          =   540
      Left            =   9180
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   90
      Width           =   1230
   End
   Begin VB.CommandButton cmdExe 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh Mail Service"
      Height          =   540
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   90
      Width           =   1230
   End
   Begin VB.CommandButton cmdCheckMail 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Check Mail Status.."
      Height          =   540
      Left            =   6255
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   90
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   540
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   90
      Width           =   1230
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete Invalid Mail.."
      Height          =   540
      Left            =   3105
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton cmdRem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear (Mail Sent List)"
      Height          =   540
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   1890
   End
   Begin VB.CommandButton cmdref 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   540
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   1095
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   8676
      Left            =   0
      TabIndex        =   0
      Top             =   852
      Width           =   18372
      _cx             =   32411
      _cy             =   15293
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
      ForeColorSel    =   -2147483647
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Frame Frame1_log 
         Caption         =   "Log File Details"
         Height          =   8565
         Left            =   0
         TabIndex        =   8
         Top             =   36
         Visible         =   0   'False
         Width           =   14550
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            Height          =   420
            Left            =   13770
            TabIndex        =   10
            Top             =   135
            Width           =   555
         End
         Begin RichTextLib.RichTextBox r1 
            Height          =   7815
            Left            =   45
            TabIndex        =   9
            Top             =   630
            Width           =   14280
            _ExtentX        =   25188
            _ExtentY        =   13780
            _Version        =   393217
            ScrollBars      =   3
            RightMargin     =   20000
            TextRTF         =   $"frmMailSt.frx":000C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F4 For Delete Invalid Mail...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   5625
      TabIndex        =   14
      Top             =   9585
      Width           =   2580
   End
   Begin VB.Label lblsendedmail 
      Height          =   255
      Left            =   2625
      TabIndex        =   4
      Top             =   9585
      Width           =   2715
   End
   Begin VB.Label lblPendingMail 
      Height          =   255
      Left            =   30
      TabIndex        =   3
      Top             =   9585
      Width           =   2415
   End
End
Attribute VB_Name = "frmMailSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vs_rs As New ADODB.Recordset

Private Sub cmdBlukMail_Click()
If MsgBox("want to send bluk mail ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "update MailDetails set MailSended='n' where mailsended='Bulk Mail...'"
   fillGrid
End If
End Sub

Private Sub cmdCheckMail_Click()

'
'For I = 1 To vs.rows - 1
'
'con.Execute "insert into MailDetails1(bill,mail,BillType,address1,HeadEmail,address4) values('" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 4) & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 5) & "','" & vs.TextMatrix(I, 6) & "','" & vs.TextMatrix(I, 8) & "')"
'
'Next
'
'
'
'Frame1_log.Visible = True
'
'If session = "2018-19" Then
'
'r1.filename = "\\192.168.0.103\Mail_System1819\bin\Debug\LogFile.txt"
'r1.LoadFile (r1.filename)
'
'ElseIf session = "2017-18" Then
'
'r1.filename = "\\192.168.0.103\Mail_System1718\bin\Debug\LogFile.txt"
'r1.LoadFile (r1.filename)
'
'End If



End Sub

Private Sub cmdClose_Click()
Frame1_log.Visible = False
End Sub

Private Sub cmdExe_Click()

'
'Dim strProgramName As String
'Dim strArgument As String
'
'If session = "2018-19" Then
'
'strProgramName = "\\192.168.0.103\Mail_System1819\bin\Debug\Mail.exe"
'strArgument = "/G"
'
'Call Shell("""" & strProgramName & """ """ & strArgument & """", vbNormalFocus)
'
'
'ElseIf session = "2017-18" Then
'
'strProgramName = "\\192.168.0.103\Mail_System1718\bin\Debug\Mail.exe"
'strArgument = "/G"
'
'Call Shell("""" & strProgramName & """ """ & strArgument & """", vbNormalFocus)
'
'
'
'End If




End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdref_Click()
fillGrid
End Sub
Private Sub cmdRem_Click()
  If vs.TextMatrix(vs.RowSel, 1) <> "" Then
  con.Execute "delete from MailDetails where MailSended='y'"
  
'  If vs.TextMatrix(vs.RowSel, 7) = "y" Then
'     CON.Execute "delete from MailDetails where Bill='" & vs.TextMatrix(vs.RowSel, 2) & "'"
'  End If
  For k1 = 1 To vs.rows - 1
  If vs.TextMatrix(k1, 2) <> "" Then
     con.Execute "delete from tempLedger7 where rptid='" & vs.TextMatrix(k1, 2) & "'"
  End If
  Next
  
  
  fillGrid
  End If
  
End Sub

Private Sub cmdRemove_Click()

If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "delete from MailDetails where id=" & vs.TextMatrix(vs.RowSel, 1) & ""
   vs.RemoveItem (vs.RowSel)
   'fillGrid
End If

End Sub



Private Sub cmdUpdateCC_Click()
frmUpdateCCMail.Show
End Sub

Private Sub Form_Load()
Me.top = 600
Me.Left = 200
fillGrid

If LCase(UserName) = "admin" Then
   cmdRemove.Visible = True
   cmdExe.Visible = True
   
Else
   'cmdRemove.Visible = False
   cmdExe.Visible = False
End If

End Sub
Sub fillGrid()
   
   If vs_rs.State = 1 Then vs_rs.close
   vs_rs.Open "select id,Bill,BillType,Mail as [Party Mail Id],RepEmail as RepEmail_CC,HeadEmail,MailSended as MailStatus,address1 as Party,Dates,Manager,WhatsappStatus from MailDetails order by id desc", con
   Set vs.DataSource = vs_rs
   
   If RS.State = 1 Then RS.close
   RS.Open "select count(*) from MailDetails where MailSended='n'", con
   If Not IsNull(RS(0)) Then
     If RS(0) > 0 Then
      lblPendingMail.Caption = "Total Painding mail : " & RS(0)
     End If
   End If
   
   If RS.State = 1 Then RS.close
   RS.Open "select count(*) from MailDetails where MailSended='y'", con
   If Not IsNull(RS(0)) Then
      If RS(0) > 0 Then
        lblsendedmail.Caption = "Total Sended mail : " & RS(0)
      End If
   End If
   
   
   DoEvents
   DoEvents
   
   vs.ColWidth(4) = 3200
   vs.ColWidth(5) = 2500
   vs.ColWidth(6) = 2200
   vs.ColWidth(7) = 1000
   vs.ColWidth(8) = 2200
   vs.ColWidth(9) = 1800
   vs.ColWidth(10) = 1500
   
   
   
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     If vs.Col = 4 Then
        con.Execute "update MailDetails set mail='" & vs.TextMatrix(vs.RowSel, 4) & "' where bill='" & vs.TextMatrix(vs.RowSel, 2) & "'"
     End If
     If vs.Col = 7 Then
        con.Execute "update MailDetails set MailSended='" & vs.TextMatrix(vs.RowSel, 7) & "' where bill='" & vs.TextMatrix(vs.RowSel, 2) & "'"
     End If
     
     SendKeys "{right}"
  End If
  
  If KeyCode = 115 Then
    If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
     con.Execute "delete from MailDetails where id=" & vs.TextMatrix(vs.RowSel, 1) & ""
     vs.RemoveItem (vs.RowSel)
    End If
  End If
  
End Sub

Private Sub vs_SelChange()
   If (vs.Col = 7 Or vs.Col = 4) Then
      vs.Editable = flexEDKbdMouse
   Else
      vs.Editable = flexEDNone
   End If
End Sub
