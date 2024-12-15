VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form issueagent 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6554.57
   ScaleMode       =   0  'User
   ScaleWidth      =   11190
   Begin VB.Frame panel 
      Caption         =   "Books Issue Statement"
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
      Height          =   5820
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   10725
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   135
         TabIndex        =   9
         Top             =   4815
         Width           =   9735
         Begin VB.CommandButton cancel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   675
            Left            =   7305
            Picture         =   "issueag.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   45
            Width           =   1155
         End
         Begin VB.CommandButton ok 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&OK"
            Enabled         =   0   'False
            Height          =   675
            Left            =   6105
            Picture         =   "issueag.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   45
            Width           =   1155
         End
         Begin VB.CommandButton Printcmd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   675
            Left            =   4905
            Picture         =   "issueag.frx":116E
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   45
            Width           =   1155
         End
         Begin VB.CommandButton search 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   675
            Left            =   3690
            Picture         =   "issueag.frx":1D52
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   45
            Width           =   1155
         End
         Begin VB.CommandButton delete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Enabled         =   0   'False
            Height          =   675
            Left            =   2490
            Picture         =   "issueag.frx":2936
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   45
            Width           =   1155
         End
         Begin VB.CommandButton Edit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   675
            Left            =   1290
            Picture         =   "issueag.frx":351A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   45
            Width           =   1155
         End
         Begin VB.CommandButton Add 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   675
            Left            =   90
            Picture         =   "issueag.frx":395C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   45
            Width           =   1155
         End
         Begin VB.CommandButton CommandQuit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   675
            Left            =   8505
            Picture         =   "issueag.frx":4540
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   45
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         Height          =   585
         Left            =   255
         TabIndex        =   4
         Top             =   7125
         Width           =   4890
         Begin VB.CommandButton Command4 
            Caption         =   "A&gent Master"
            Height          =   300
            Left            =   105
            TabIndex        =   8
            Top             =   195
            Width           =   1155
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&District Master"
            Height          =   300
            Left            =   1245
            TabIndex        =   7
            Top             =   195
            Width           =   1185
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&City Master"
            Height          =   300
            Left            =   2415
            TabIndex        =   6
            Top             =   195
            Width           =   1080
         End
         Begin VB.CommandButton Command3 
            Caption         =   "College &Master"
            Height          =   300
            Left            =   3480
            TabIndex        =   5
            Top             =   195
            Width           =   1260
         End
      End
      Begin VB.TextBox txtSearcha 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7275
         TabIndex        =   3
         Top             =   405
         Width           =   2895
      End
      Begin VB.TextBox txtsearchs 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2670
         TabIndex        =   2
         Top             =   435
         Width           =   3390
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   9990
         Top             =   4860
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   9945
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5355
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Bindings        =   "issueag.frx":5124
         Height          =   3240
         Left            =   135
         TabIndex        =   1
         Top             =   855
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5715
         _Version        =   393216
         BackColorFixed  =   7917545
         BackColorBkg    =   16777215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid11 
         Height          =   570
         Left            =   135
         TabIndex        =   18
         Top             =   990
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   1005
         _Version        =   393216
         BackColor       =   -2147483644
         BackColorFixed  =   12632256
         BackColorBkg    =   -2147483645
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLineWidthBand=   1
      End
      Begin MSMask.MaskEdBox date1 
         Height          =   315
         Left            =   2385
         TabIndex        =   19
         Top             =   2925
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   915
         Left            =   90
         Top             =   4770
         Width           =   9825
      End
      Begin VB.Label Label7 
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   25
         Top             =   4140
         Width           =   555
      End
      Begin VB.Label total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0078CFE9&
         Height          =   285
         Left            =   8250
         TabIndex        =   24
         Top             =   4140
         Width           =   1605
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press F2 for Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   2790
         TabIndex        =   23
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label lblSearch 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6105
         TabIndex        =   22
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label lblSc 
         BackStyle       =   0  'Transparent
         Caption         =   "Schoo/College Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   765
         TabIndex        =   21
         Top             =   465
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select row to delete a record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   20
         Top             =   4185
         Width           =   2925
      End
   End
End
Attribute VB_Name = "issueagent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mrpeat As Boolean
Public orderchk As Boolean
Public partchk As Boolean
Public bindchk As Boolean
Public gridchk As Boolean
Public mode As String
Dim vn As Integer
Dim flag As Boolean
Dim set1 As Boolean
Dim vflag As Boolean
Dim sgot As Boolean
Dim RS As New ADODB.Recordset
Dim cityflag As Boolean
Dim cityname As String
Sub try()
gridchk = True
Grid1.Row = 1
Grid1.Col = 8
If Grid1.Text = "" Then
gridchk = False
Grid1.SetFocus
End If
End Sub
Sub grid_ini()
 'Me.Grid1.Clear
Me.Grid1.Cols = 8
Me.Grid1.Rows = 2
Me.Grid1.RowHeight(0) = 400
Me.Grid1.RowHeight(1) = 300
    Me.Grid1.ColWidth(0) = 600
    Me.Grid1.ColWidth(1) = 1700
    Me.Grid1.ColWidth(2) = 750
    Me.Grid1.ColWidth(3) = 1150
    Me.Grid1.ColWidth(4) = 1250
    Me.Grid1.ColWidth(5) = 1250
    Me.Grid1.ColWidth(6) = 2200
    Me.Grid1.ColWidth(7) = 800
    Me.Grid1.TextMatrix(0, 0) = ""

    Me.Grid1.TextMatrix(0, 1) = "AGENT NAME"
    Me.Grid1.TextMatrix(0, 2) = "VNO"
    Me.Grid1.TextMatrix(0, 3) = "DATE"
    Me.Grid1.TextMatrix(0, 4) = "DISTRICT"
    Me.Grid1.TextMatrix(0, 5) = "CITY"
    Me.Grid1.TextMatrix(0, 6) = "COLLEGE NAME"
    Me.Grid1.TextMatrix(0, 7) = "QTY"

''    Me.Grid1.TextMatrix(0, 1) = "VNO"
''    Me.Grid1.TextMatrix(0, 2) = "DATE"
''    Me.Grid1.TextMatrix(0, 3) = "AGENT NAME"
''    Me.Grid1.TextMatrix(0, 4) = "DISTRICT"
''    Me.Grid1.TextMatrix(0, 5) = "CITY"
''    Me.Grid1.TextMatrix(0, 6) = "COLLEGE NAME"
''    Me.Grid1.TextMatrix(0, 7) = "QTY"
End Sub
Sub adddisab()
'Me.Edit.Enabled = False
'Me.mvfrst.Enabled = False
'Me.Mvlst.Enabled = False
'Me.Mvnxt.Enabled = False
'Me.Mvprv.Enabled = False
'Me.search.Enabled = False
Me.Printcmd.Enabled = False
End Sub
Sub addenab()
Me.Edit.Enabled = True
'Me.mvfrst.Enabled = True
'Me.Mvlst.Enabled = True
'Me.Mvnxt.Enabled = True
'Me.Mvprv.Enabled = True
Me.search.Enabled = True
Me.Printcmd.Enabled = True
End Sub
Private Sub Add_Click()

Clearvalue
ok.Visible = True
ok.Caption = "&OK"

'lblSc.Enabled = False
'lblSearch.Visible = False
Me.txtSearcha.Enabled = False
Me.txtsearchs.Enabled = False
Me.delete.Enabled = False

mode = ""
pstno = 0
adddisab
Grid1.Enabled = True
Dim bill As ADODB.Recordset
Set bill = New ADODB.Recordset
Me.ok.Enabled = True
mode = "add"
Me.Grid1.Col = 1
Me.Grid1.Row = 1
Me.Grid1.SetFocus
SendKeys "{f2}"
'popuplist12 "Select agentname from agentmaster order by agentname", CON

'Me.bill_no.SetFocus
SetButton Add, Edit, ok, delete

End Sub
Private Sub Add_GotFocus()
Me.delete.Enabled = False
End Sub

Private Sub cancel_Click()
Clearvalue
'Add_Click
'Progress.Show
lastrecord

txtSearcha.Enabled = False
txtsearchs.Enabled = False
'Me.lblSearch.Enabled = False
'Me.lblSc.Enabled = False

Me.ok.Enabled = False
Me.ok.Caption = "&OK"
'Grid1.Enabled = False
addenab
'Mvlst_Click
End Sub
Private Sub updateRecord()
Dim RS As New ADODB.Recordset
Dim I, X As Integer
Dim update As Boolean
'Me.Grid1.Rows = Me.Grid1.Rows - 1
For I = 1 To Grid1.Rows - 1
If RS.State = 1 Then RS.close
RS.Open "Select * from info where " & stringyear & " and vno='" & UCase(Grid1.TextMatrix(I, 2)) & "' and aname= '" & Trim(Grid1.TextMatrix(I, 1)) & "'", con, adOpenDynamic, adLockPessimistic

      If RS.RecordCount > 0 Then
          '  RS("vno") = Grid1.TextMatrix(I, 1) & ""
            RS("date") = CDate(Grid1.TextMatrix(I, 3))
            RS("aname") = Grid1.TextMatrix(I, 1) & ""
            RS("district") = Grid1.TextMatrix(I, 4) & ""
            RS("city") = Grid1.TextMatrix(I, 5) & ""
            RS("scname") = Grid1.TextMatrix(I, 6) & ""
            RS("qty") = Grid1.TextMatrix(I, 7) & ""
            RS.update
        update = True
        End If
    Next I
    If update = True Then
    MsgBox "The Record Updated Successfully", vbInformation
    End If
   Me.ok.Enabled = False
   'End If
  RS.close
'Me.ok.Enabled = False
Me.Grid1.Enabled = False
addenab
Me.Add.SetFocus
End Sub
Private Sub cmbSearch_Click()
Select Case cmbSearch.Text
 Case "By Voucher No."
         grid_ini
         lblsc.Visible = False
         lblv.Visible = True
         lblSearch.Visible = False
         txtsearch.Visible = True
         txtsearch.Text = ""
         txtSearcha.Visible = False
         txtsearchs.Visible = False
 Case "By Agent Name"
         grid_ini
         lblsc.Visible = False
         lblv.Visible = False
         lblSearch.Visible = True
         txtSearcha.Visible = True
         txtSearcha.Text = ""
         txtsearchs.Visible = False
         txtsearch.Visible = False
 Case "By School/College Name"
         grid_ini
         lblsc.Visible = True
     '    lblv.Visible = False
         lblSearch.Visible = False
         txtsearchs.Visible = True
         txtsearchs.Text = ""
         'txtSearch.Visible = False
         txtSearcha.Visible = False
 End Select
End Sub

Private Sub cancel_GotFocus()
Me.delete.Enabled = False
End Sub

Private Sub Command1_Click()
bookmaster.SSTab1.Tab = 2
bookmaster.Show
End Sub

Private Sub Command2_Click()
'Me.WindowState = 2
bookmaster.SSTab1.Tab = 5
bookmaster.Show

End Sub

Private Sub Command3_Click()
FrmSchool.Show
End Sub
Private Sub Command4_Click()
bookmaster.SSTab1.Tab = 3
bookmaster.Show
bookmaster.Commandmasteradd.SetFocus
End Sub

Private Sub CommandQuit_GotFocus()
Me.delete.Enabled = False
End Sub
Private Sub date_LostFocus()
Grid1.Col = 3
Grid1.Text = date1.Text
End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Grid1.Col = 3

'If CDate(date1.Text) = False Then
'MsgBox "Pls enter correct date format"
'date1.SetFocus
'Else
Grid1.Text = date1.Text
date1.Text = "__/__/____"
Grid1.SetFocus

'popuplist1.Show
popuplistModel.Text1.Text = ""

'SendKeys "{f2}"
'popuplist12 "Select agentname from agentmaster order by agentname", CON
Grid1.Col = 4
Grid1_GotFocus
date1.Visible = False
'End If
End If
End Sub

Private Sub date1_LostFocus()
date1.Visible = False
End Sub

Private Sub Delete_Click()



If Grid1.Rows = 2 Then
Grid1.Rows = 3
End If
Dim vn As String
If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If Grid1.Row > 0 Then
                
           con.Execute "delete  from info where " & stringyear & " and vno='" & Grid1.TextMatrix(Grid1.RowSel, 2) & "' and aname= '" & Grid1.TextMatrix(Grid1.RowSel, 1) & "'"
           Grid1.RemoveItem Grid1.Row
           
           Else
      MsgBox "You cannnot delete first and fixed row you can edit it"
    End If
End If

End Sub

Private Sub delete_GotFocus()
Me.delete.Enabled = True
End Sub
Private Sub delete_LostFocus()
Me.delete.Enabled = False
End Sub

Private Sub Edit_Click()
'grid_ini



ok.Visible = True
ok.Enabled = True
ok.Caption = "&Update"

txtSearcha.Enabled = False
txtsearchs.Enabled = False
Printcmd.Enabled = True

Me.delete.Enabled = False
Me.ok.Enabled = True
mode = "edit"
Grid1.Enabled = True

Grid1.Row = 1
Grid1.Col = 2


SetButton Add, Edit, ok, delete

End Sub

Private Sub Edit_GotFocus()
Me.delete.Enabled = False
End Sub
Private Sub Form_Activate()
'popuplist11 "Select DISTRICTNAME from DISTRICTS order by DISTRICTNAME ", CON
ok.Enabled = False
delete.Enabled = False
End Sub

Private Sub Form_GotFocus()
Me.delete.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()

Me.Top = 200
Me.Left = 200

'Data1.ConnectionString = CON
'Data1.DatabaseName = VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
 
Label10.Visible = False
'grid_ini
'popuplist11 "Select college, collegeid,cityid from college order by college ", CON
'popuplist11 "Select DISTRICTNAME from DISTRICTS order by DISTRICTNAME ", CON

'popuplist1 "Select DISTRICTNAME from DISTRICTS order by DISTRICTNAME ", CON

lblsc.Enabled = False
lblSearch.Enabled = False
txtSearcha.Enabled = False
txtsearchs.Enabled = False


lastrecord

BackColorFrom Me

DoEvents
DoEvents
delete.Enabled = False
ok.Enabled = False
DoEvents
DoEvents


End Sub
Private Sub fillcmb()
 cmbSearch.AddItem "By Voucher No."
 cmbSearch.AddItem "By Agent Name"
 cmbSearch.AddItem "By School/College Name"
End Sub


Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub Grid1_Click()

On Error Resume Next

If mode = "edit" And Grid1.Col = 1 Then
Grid1.Col = 2
Grid1.SetFocus
date1.Visible = True
date1.ZOrder (1)
date1.Text = Grid1.Text
date1.Width = Grid1.ColWidth(Grid1.Col)
date1.Height = Grid1.CellHeight
date1.Left = Grid1.CellLeft + 100
date1.Top = Grid1.Top + Grid1.CellTop '- 50
date1.SetFocus
date1.ZOrder (0)
End If

If mode = "edit" And Grid1.Col = 3 Then
If Grid1.Col = 3 And date1.Text <> "" Then
date1.Visible = True
date1.ZOrder (1)
'date1.Text = "12/05/2005"
date1.Text = Grid1.Text
date1.Width = Grid1.ColWidth(Grid1.Col)
date1.Height = Grid1.CellHeight
date1.Left = Grid1.CellLeft + 100
date1.Top = Grid1.Top + Grid1.CellTop '- 50
date1.SetFocus
date1.ZOrder (0)
End If
End If


If mode = "add" And Grid1.Col = 3 Then
If Grid1.Col = 3 And date1.Text <> "" Then
date1.Visible = True
date1.ZOrder (1)
date1.Text = Grid1.Text
date1.Width = Grid1.ColWidth(Grid1.Col)
date1.Height = Grid1.CellHeight
date1.Left = Grid1.CellLeft + 100
date1.Top = Grid1.Top + Grid1.CellTop '- 50
date1.SetFocus
date1.ZOrder (0)
End If
End If

SetButton Commandadd, Commandedit, Commandsave, Commanddelete

End Sub

Private Sub Grid1_GotFocus()
'Dim RSK As ADODB.Recordset
'Set RSK = New ADODB.Recordset

    If PopUpValue1 <> "" Then
    Grid1.Text = PopUpValue1
    End If



If Me.Grid1.Col = 4 Then
End If

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""


Me.delete.Enabled = True


End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

Dim RSK As ADODB.Recordset
Set RSK = New ADODB.Recordset

Dim ST, r As Integer
r = Grid1.Row
Grid1.TextMatrix(1, 0) = 1
If Grid1.Col = 2 Then
    If Grid1.Text = "" And KeyCode = 13 Then
          ok.SetFocus
          Exit Sub
    End If
End If



If Me.Grid1.Col = 1 And KeyCode = 113 Then
   popuplistModel10 "Select agentname from agentmaster where " & stringyear & " order by agentname", con
End If


If (Grid1.Col = 2 And KeyCode = 13) And (Grid1.Text = "" Or Grid1.Text = "0") Then
Me.Grid1.SetFocus
Grid1.Col = 2
Exit Sub
End If



If Me.Grid1.Col = 3 And KeyCode = 13 Then
'popuplist1.Show
popuplistModel.Text1.Text = ""
End If


If Me.Grid1.Col = 4 And KeyCode = 113 Then
popuplistModel10 "Select DISTRICTNAME from DISTRICTS order by DISTRICTNAME ", con
popuplistModel.Text1.Text = ""
End If


If Me.Grid1.Col = 4 And KeyCode = 13 Then
   popuplistModel10 "Select city from city where " & stringyear & " and district = '" + Grid1.TextMatrix(r, 4) + "' order by city", con
End If


If Me.Grid1.Col = 5 And KeyCode = 113 Then
popuplistModel10 "Select city from city where " & stringyear & " and district = '" + Grid1.TextMatrix(r, 4) + "' order by city", con
End If


If Me.Grid1.Col = 5 And KeyCode = 13 Then
If RS.State = 1 Then RS.close
RS.Open "select cityid from city where " & stringyear & " and city = '" + Grid1.TextMatrix(r, 5) + "' and district= '" + Grid1.TextMatrix(r, 4) + "'", con, adOpenStatic, adLockReadOnly
If RS.RecordCount > 0 Then
a = RS!cityId
RS.close
Else
RS.close
End If
popuplistModel10 "Select college from college where " & stringyear & " and cityid = " & a, con

End If

If Me.Grid1.Col = 6 And KeyCode = 113 Then
popuplistModel10 "Select college from college where " & stringyear & " and city = '" + Grid1.TextMatrix(r, 5) + "' order by college", con
End If


    If Grid1.Col = 2 And KeyCode = 13 Then
    Dim voucher As String
    Dim Agent As String
    voucher = Grid1.TextMatrix(Grid1.Row, 2)
    Agent = Grid1.TextMatrix(Grid1.Row, 1)
    For I = 1 To Grid1.Rows - 2
    If Grid1.TextMatrix(I, 1) = Agent And Grid1.TextMatrix(I, 2) = voucher Then
    MsgBox "voucher no cannot be repeated"
    Grid1.Col = 2
    Grid1.SetFocus
    'vflag = True
    Exit Sub
    End If
    Next
    End If


            If KeyCode = 115 Then
                    If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
                        If Grid1.Row > 1 Then
                            Grid1.RemoveItem Grid1.Row
                        Else
                            MsgBox "You cannnot delete first and fixed row you can edit it"
                        End If
                    End If
            End If

If Grid1.Col = 2 And KeyCode = 13 Then
      If mode = "edit" Then

                        Grid1.Col = 3
                        Grid1.SetFocus
                        date1.Visible = True
                        date1.ZOrder (1)
                        date1.Text = Grid1.Text
                        date1.Width = Grid1.ColWidth(Grid1.Col)
                        date1.Height = Grid1.CellHeight
                        date1.Left = Grid1.CellLeft + 100
                        date1.Top = Grid1.Top + Grid1.CellTop '- 50
                        date1.SetFocus
                        date1.ZOrder (0)

'         Grid1.col = 2
 '          Grid1.SetFocus
   '        Exit Sub
       Else

                Grid1.Row = r
                If RSK.State = 1 Then RSK.close
                RSK.Open "Select * from INFO where " & stringyear & " and vno = '" & Trim(Grid1.TextMatrix(r, 2)) & "' and aname= '" & Trim(Grid1.TextMatrix(r, 1)) & "'", con, adOpenStatic, adLockPessimistic
                If RSK.RecordCount > 0 Then

                    MsgBox ("This Record is Already Define....?")

                      Grid1.Row = Grid1.Row
                       Grid1.Col = 2
                       Grid1.Text = ""
                        Grid1.SetFocus
                        Exit Sub
                Else

                    Grid1.Col = 3
                    Grid1.SetFocus
                    date1.Visible = True
                    date1.ZOrder (1)
                    'date1.Text = Grid1.Text
                    date1.Width = Grid1.ColWidth(Grid1.Col)
                    date1.Height = Grid1.CellHeight
                    date1.Left = Grid1.CellLeft + 100
                    date1.Top = Grid1.Top + Grid1.CellTop '- 50
                    date1.SetFocus
                    date1.ZOrder (0)
                    Dim Str As String
                    Str = Grid1.Text
                    Grid1.Text = UCase(Str)
                    Exit Sub
                End If

End If
End If

If Grid1.Col = 1 And KeyCode = 13 Then Grid1.Col = 2: Exit Sub

'If vflag = False And KeyCode = 13 Then Grid1.col = 3: Exit Sub
'If grid1.col = 2 And KeyCode = 13 Then grid1.col = 3: Exit Sub
If Grid1.Col = 3 And KeyCode = 13 Then Grid1.Col = 4: Exit Sub
If Grid1.Col = 4 And KeyCode = 13 Then Grid1.Col = 5: Exit Sub
If Grid1.Col = 5 And KeyCode = 13 Then Grid1.Col = 6: Exit Sub
If Grid1.Col = 6 And KeyCode = 13 Then Grid1.Col = 7: Exit Sub

If Grid1.Col = 7 And KeyCode = 13 And Grid1.Text = "" Then
MsgBox "Please enter some value"
Me.Grid1.SetFocus
Grid1.Col = 7
Exit Sub
End If

If Grid1.Col = 7 And KeyCode = 13 Then
calc
    
    SaveSingleLine
    If Grid1.Row = Grid1.Rows - 1 Then
    Grid1.AddItem "", Grid1.Row + 1

    Grid1.Row = Grid1.Row + 1
    Grid1.TextMatrix(r + 1, 0) = r + 1


    Grid1.TextMatrix(r + 1, 1) = Grid1.TextMatrix(r, 1)
    Grid1.TextMatrix(r + 1, 4) = Grid1.TextMatrix(r, 4)
   ' Grid1.TextMatrix(r + 1, 2) = Grid1.TextMatrix(r, 2)



    End If
    
    ST = Me.Grid1.Row
    Me.Grid1.RowHeight(ST) = 300
    SendKeys "{DOWN}"
    For I = 0 To 14
    SendKeys "{LEFT}"
    Next I

Grid1.Col = 1
Grid1.SetFocus
'SendKeys "{f2}"
'popuplist12 "Select agentname from agentmaster order by agentname", CON

End If

Me.delete.Enabled = True



End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
Me.delete.Enabled = True

If Grid1.Col = 2 And mode = "add" Then
        If KeyAscii = 8 Then
        If Len(Trim(Grid1.Text)) <> 0 Then
                Grid1.Text = Left(Grid1.Text, (Len(Grid1.Text) - 1))
        End If
            ElseIf (KeyAscii <> 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
            Grid1.Text = Grid1.Text + Chr(KeyAscii)
        End If
End If



If Grid1.Col = 3 Then
    If Len(Trim(Grid1.Text)) = 80 Then
    KeyAscii = 0
    End If
        If KeyAscii = 8 Then
           If Len(Trim(Grid1.Text)) <> 0 Then
                Grid1.Text = Left(Grid1.Text, (Len(Grid1.Text) - 1))
           End If
           ElseIf (KeyAscii <> 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
            Grid1.Text = Grid1.Text + Chr(KeyAscii)
        End If
End If

If Grid1.Col = 1 Then
 '       If KeyAscii = 8 Then
' '       If Len(Trim(grid1.Text)) <> 0 Then
  '              grid1.Text = Left(grid1.Text, (Len(grid1.Text) - 1))
   '     End If
    '        ElseIf (KeyAscii <> 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
     '       grid1.Text = grid1.Text + Chr(KeyAscii)
        'End If
End If

If Grid1.Col = 4 Then
        'If KeyAscii = 8 Then
        'If Len(Trim(grid1.Text)) <> 0 Then
        '        grid1.Text = Left(grid1.Text, (Len(grid1.Text) - 1))
        'End If
        '    ElseIf (KeyAscii <> 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
        '    grid1.Text = grid1.Text + Chr(KeyAscii)
        'End If
End If
If Grid1.Col = 5 Then
'        If KeyAscii = 8 Then
'        If Len(Trim(grid1.Text)) <> 0 Then
'                grid1.Text = Left(grid1.Text, (Len(grid1.Text) - 1))
'        End If
'            ElseIf (KeyAscii <> 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
'            grid1.Text = grid1.Text + Chr(KeyAscii)
'        End If
End If

If Grid1.Col = 6 Then
'        If KeyAscii = 8 Then
'        If Len(Trim(grid1.Text)) <> 0 Then
'                grid1.Text = Left(grid1.Text, (Len(grid1.Text) - 1))
'        End If
'            ElseIf (KeyAscii <> 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
'            grid1.Text = grid1.Text + Chr(KeyAscii)
'        End If
End If


'Rate Amount quantity coding

If Grid1.Col = 7 Then
    If Len(Trim(Grid1.Text)) = 10 Then
    KeyAscii = 0
    End If
    If KeyAscii = 8 Then
        If Len(Trim(Grid1.Text)) <> 0 Then
                Grid1.Text = Left(Grid1.Text, (Len(Grid1.Text) - 1))
        End If
    End If
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
    If KeyAscii <> 8 Then
        Grid1.Text = Grid1.Text + Chr(KeyAscii)
    End If
End If






End Sub
Private Sub Grid1_LostFocus()
'Me.delete.Enabled = False
End Sub
Private Sub ok_Click()



If ok.Caption = "&OK" Then
 saverecord
Else
 updateRecord
End If
End Sub
Sub SaveSingleLine()
Dim RS As New ADODB.Recordset
Dim I As Integer
 RS.Open "select * from Info where " & stringyear & " and aname='" & Grid1.TextMatrix(Grid1.RowSel, 1) & "' and vno='" & Grid1.TextMatrix(Grid1.RowSel, 2) & "'", con, adOpenStatic, adLockPessimistic
 If RS.EOF = False Then
    MsgBox " V. Number  is Already Exist ..", vbCritical
    Exit Sub
 End If
RS.AddNew
RS("vno") = UCase(Grid1.TextMatrix(Grid1.RowSel, 2))
RS("date") = CDate(Grid1.TextMatrix(Grid1.RowSel, 3))
RS("aname") = Grid1.TextMatrix(Grid1.RowSel, 1)
RS("district") = Grid1.TextMatrix(Grid1.RowSel, 4)
RS("city") = Grid1.TextMatrix(Grid1.RowSel, 5)
RS("scname") = Grid1.TextMatrix(Grid1.RowSel, 6)
RS("qty") = Val(Grid1.TextMatrix(Grid1.RowSel, 7))
RS.update
RS.close
End Sub

Private Sub saverecord()
Dim RS As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim I As Integer
 RS.Open "Info", con, adOpenStatic, adLockPessimistic
 Me.Grid1.Rows = Me.Grid1.Rows - 1
   If RS.RecordCount > 0 Then
     For I = 1 To Grid1.Rows - 1
       RS.MoveLast
     If rs1.State = 1 Then rs1.close
     rs1.Open "select * from Info where " & stringyear & " and vno='" & UCase(Grid1.TextMatrix(I, 2)) & "' and convert(smalldatetime,date,103)=convert(smalldatetime,'" & Grid1.TextMatrix(I, 3) & "',103) and aname='" & Grid1.TextMatrix(I, 1) & "'", con, adOpenDynamic, adLockOptimistic
     
     
     If rs1.EOF = True Then
       RS.AddNew
            RS("vno") = UCase(Grid1.TextMatrix(I, 2))
            RS("date") = CDate(Grid1.TextMatrix(I, 3))
            RS("aname") = Grid1.TextMatrix(I, 1)
            RS("district") = Grid1.TextMatrix(I, 4)
            RS("city") = Grid1.TextMatrix(I, 5)
            RS("scname") = Grid1.TextMatrix(I, 6)
            RS("qty") = Val(Grid1.TextMatrix(I, 7))
      RS.update
      End If
      RS.MoveNext
     Next I
     MsgBox "The Record Saved Successfully", vbInformation
    End If
    RS.close
Me.ok.Enabled = False
Me.Grid1.Enabled = False
addenab
Me.Add.SetFocus
End Sub

Private Sub order_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   TextPaperSize.SetFocus
End If
End Sub
Private Sub order_no_LostFocus()
 'Binder_id.SetFocus
End Sub
Private Sub ok_GotFocus()
Me.delete.Enabled = False
End Sub


Private Sub commandQuit_Click()
Unload Me
End Sub



Private Sub Printcmd_GotFocus()
'Me.delete.Enabled = False
'SetButton Commandadd, Commandedit, Commandsave, Commanddelete

End Sub
Private Sub search_Click()
grid_ini


         lblSearch.Enabled = True
         txtSearcha.Enabled = True
         lblsc.Enabled = True
         txtsearchs.Enabled = True

         txtSearcha.Text = ""
         txtsearchs.Text = ""
         txtsearchs.SetFocus
'Me.Grid1.Enabled = False
Me.ok.Enabled = False
End Sub
Private Sub Textfirmname_GotFocus()
'mode = ""
'addenab
'If PopUpValue1 <> "" Then
'   Textfirmname.Text = PopUpValue1
'   Textfirmid.Text = PopUpValue2
'   firmpictfilename.Text = PopUpValue3

'Me.bill_no.SetFocus
'Mvlst_Click
'Me.Add.SetFocus

'End If
'PopUpValue1 = ""
'PopUpValue2 = ""
'PopUpValue3 = ""
End Sub
Sub Clearvalue()
Grid1.Clear
grid_ini
total.Caption = ""
Me.txtSearcha = ""
Me.txtsearchs = ""
End Sub


Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
  txtSearch_LostFocus
 End If
End Sub

'Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 13 Then
'   txtSearch_LostFocus
'  End If
'End Sub

Private Sub txtSearch_LostFocus()
'Dim rs As New ADODB.Recordset
Dim I As Integer
I = 1
RS.Open "Select * from info where " & stringyear & " and vno = '" & Trim(txtsearch) & "'", con, adOpenStatic, adLockReadOnly
 If RS.RecordCount > 0 Then
' For i = 1 To Grid1.Rows
  Me.Grid1.TextMatrix(I, 0) = I
  Me.Grid1.TextMatrix(I, 2) = RS("vno") & ""
  Me.Grid1.TextMatrix(I, 3) = RS("date") & ""
  Me.Grid1.TextMatrix(I, 1) = RS("aname") & ""
  Me.Grid1.TextMatrix(I, 4) = RS("district") & ""
  Me.Grid1.TextMatrix(I, 5) = RS("scname") & ""
  Me.Grid1.TextMatrix(I, 6) = RS("city") & ""
  Me.Grid1.TextMatrix(I, 7) = RS("qty") & ""
 RS.MoveNext
 'Next i
 Else
  MsgBox ("Record not found,Please Check the value"), vbInformation
  grid_ini
 End If
 RS.close
End Sub

Private Sub search_GotFocus()
Me.delete.Enabled = False
End Sub

Private Sub Timer1_Timer()
If set1 Then
Label10.Visible = Not Label10.Visible
Label10.ForeColor = RGB(255, 255, 255) * Rnd()
End If
End Sub

Private Sub txtSearcha_GotFocus()
'Me.Clearvalue
    If PopUpValue1 <> "" Then
    txtSearcha = PopUpValue1
    Label10.Visible = True
    set1 = True
    End If
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
Label10.Visible = True
set1 = True
Me.delete.Enabled = False
End Sub

Private Sub txtSearcha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
Me.Clearvalue
popuplistModel10 "Select distinct agentname from agentmaster where " & stringyear & " order by agentname", con
End If

  If KeyCode = 13 Then
  Me.Grid1.Clear
  Me.grid_ini
  Searcha
  End If
End Sub
Sub Searcha()
If Trim(txtSearcha) <> "" Then
    Dim RS As New ADODB.Recordset
    Dim I As Integer
    I = 1
    RS.Open "Select * from info where " & stringyear & " and aname = '" & Trim(txtSearcha) & "' order by SUBSTRING(vno,1,1),CONVERT(INT, vno)", con, adOpenStatic, adLockReadOnly

    If RS.RecordCount > 0 Then
        'Progress.Show
        grid_ini
        ' For i = 1 To Grid1.Rows
        'Progress.pb1 = 0
        'Progress.pb1.max = rs.RecordCount
        Do While Not RS.EOF
            Me.Grid1.TextMatrix(I, 0) = I
            Me.Grid1.TextMatrix(I, 2) = RS("vno") & ""
            Me.Grid1.TextMatrix(I, 3) = RS("date") & ""
            Me.Grid1.TextMatrix(I, 1) = RS("aname") & ""
            Me.Grid1.TextMatrix(I, 4) = RS("district") & ""
            Me.Grid1.TextMatrix(I, 5) = RS("city") & ""
            Me.Grid1.TextMatrix(I, 6) = RS("scname") & ""
            Me.Grid1.TextMatrix(I, 7) = RS("qty") & ""
            If Not RS.EOF Then
                RS.MoveNext
                I = I + 1
                'Progress.pb1 = Progress.pb1 + 1
                'If Progress.pb1 = Progress.pb1.max Then Unload Progress
                Grid1.Rows = Grid1.Rows + 1
                calc
            End If
        Loop
        Grid1.Rows = Grid1.Rows - 1
    Else
        MsgBox ("Record not found,Please Check the value"), vbInformation
        grid_ini
        Me.txtSearcha = ""
        'rs.close
        'Exit Sub
    End If
    '' RS.close
End If
''Label10.Visible = False
''set1 = False
End Sub

Private Sub txtsearchs_GotFocus()
'Unload popuplist
Dim colname As String

Me.Clearvalue

If PopUpValue1 <> "" Then

        '''If cityflag = True Then
        txtsearchs = PopUpValue1
      
End If

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
'colname = ""
Label10.Visible = True
set1 = True
Me.delete.Enabled = False
End Sub

Private Sub txtsearchs_KeyDown(KeyCode As Integer, Shift As Integer)
Dim value As String
If KeyCode = 113 Then

cityflag = False


'''searchType = "inv"
popuplist_client "Select scname as college,City,AgentName from school  where " & stringyear & "  order by scname,City,AgentName", CCON


End If
 If KeyCode = 13 Then
   txtsearchs_LostFocus
 End If
End Sub

Private Sub txtsearchs_LostFocus()
If txtsearchs <> "" Then
Dim RS As New ADODB.Recordset
Dim I As Integer
I = 1

RS.Open "Select * from info where " & stringyear & " and scname = '" & Trim(txtsearchs) & "'", con, adOpenStatic, adLockReadOnly
 If RS.RecordCount > 0 Then
grid_ini


Do While Not RS.EOF
  Me.Grid1.TextMatrix(I, 0) = I
  Me.Grid1.TextMatrix(I, 2) = RS("vno") & ""
  Me.Grid1.TextMatrix(I, 3) = RS("date") & ""
  Me.Grid1.TextMatrix(I, 1) = RS("aname") & ""
  Me.Grid1.TextMatrix(I, 4) = RS("district") & ""
  Me.Grid1.TextMatrix(I, 5) = RS("city") & ""
  Me.Grid1.TextMatrix(I, 6) = RS("scname") & ""
  Me.Grid1.TextMatrix(I, 7) = RS("qty") & ""
         If Not RS.EOF Then
            RS.MoveNext
            I = I + 1
            Grid1.Rows = Grid1.Rows + 1
calc
        End If
Loop
 Else
  MsgBox ("Record not found,Please Check the value"), vbInformation
   grid_ini
    Me.txtsearchs = ""
  End If
 RS.close
 End If
 Label10.Visible = False
set1 = False
End Sub

Private Sub txtvno_LostFocus()
Dim RS As New ADODB.Recordset
Dim I As Integer
I = 1
RS.Open "Select * from info where " & stringyear & " and vno = '" & Trim(txtvno) & "'", con, adOpenStatic, adLockReadOnly
 If RS.RecordCount > 0 Then
  Me.Grid1.TextMatrix(I, 0) = I
  Me.Grid1.TextMatrix(I, 2) = RS("vno") & ""
  Me.Grid1.TextMatrix(I, 3) = RS("date") & ""
  Me.Grid1.TextMatrix(I, 1) = RS("aname") & ""
  Me.Grid1.TextMatrix(I, 4) = RS("district") & ""
  Me.Grid1.TextMatrix(I, 5) = RS("city") & ""
  Me.Grid1.TextMatrix(I, 6) = RS("scname") & ""
  Me.Grid1.TextMatrix(I, 7) = RS("qty") & ""
 RS.MoveNext
 Else
  MsgBox ("Record not found,Please Check the value"), vbInformation
  grid_ini
 End If
 RS.close
End Sub

Sub calc()
Dim amount As Double
For I = 1 To Grid1.Rows - 1
amount = amount + Val(Grid1.TextMatrix(I, 7))
Next I
total.Caption = amount
End Sub
Sub lastrecord()
Dim V As Integer
'''grid_ini
'''ssql = "select vno, date, aname, district, city, scname ,qty from info"
'''Dim Grs As ADODB.Recordset
'''Set Grs = New Recordset
'''  Grs.Open ssql, CON, adOpenStatic, adLockReadOnly
'''
'''
'''
'''     If Grs.RecordCount > 0 Then
'''
'''        Grid1.row = 1
'''        grid_ini
'''            Progress.PB1 = 0
'''            Progress.PB1.Max = Grs.RecordCount
'''            Do While Not Grs.EOF
'''
'''              Me.Grid1.TextMatrix(Grid1.row, 0) = Grid1.row
'''              Me.Grid1.TextMatrix(Grid1.row, 2) = Grs("vno") & ""
'''              Me.Grid1.TextMatrix(Grid1.row, 3) = Grs("date") & ""
'''              Me.Grid1.TextMatrix(Grid1.row, 1) = Grs("aname") & ""
'''              Me.Grid1.TextMatrix(Grid1.row, 4) = Grs("district") & ""
'''              Me.Grid1.TextMatrix(Grid1.row, 5) = Grs("city") & ""
'''              Me.Grid1.TextMatrix(Grid1.row, 6) = Grs("scname") & ""
'''              Me.Grid1.TextMatrix(Grid1.row, 7) = Grs("qty") & ""
'''
'''             If Not Grs.EOF Then
'''             Grs.MoveNext
'''             Progress.PB1 = Progress.PB1 + 1
'''             If Progress.PB1 = Progress.PB1.Max Then Unload Progress
'''             Grid1.Rows = Grid1.Rows + 1
'''             Grid1.row = Grid1.row + 1
'''             calc
'''             End If
'''            Loop
'''            Grid1.Rows = Grid1.Rows - 1
'''            Grid1.Enabled = True
'''     End If
'''

grid_ini
Data1.RecordSource = "select aname as Agent, Date, Vno, District, City, scname as School ,Qty from info"
Data1.Refresh

Grid1.Row = 1
V = Grid1.Rows
 
    For I = 1 To (V - 2)
        Me.Grid1.TextMatrix(Grid1.Row, 0) = Grid1.Row
        Grid1.Row = Grid1.Row + 1
    Next
             calc
End Sub
