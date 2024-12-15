VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTeacherDetail 
   ClientHeight    =   9396
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9396
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame panel 
      Height          =   9060
      Left            =   630
      TabIndex        =   12
      Top             =   180
      Width           =   13335
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2355
         TabIndex        =   2
         Top             =   1245
         Width           =   5085
      End
      Begin VB.TextBox txtaddress1 
         Appearance      =   0  'Flat
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
         Left            =   2355
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   1590
         Width           =   5100
      End
      Begin VB.TextBox txtphone 
         Appearance      =   0  'Flat
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
         Left            =   2355
         TabIndex        =   7
         Top             =   3030
         Width           =   4155
      End
      Begin VB.TextBox txtaddress2 
         Appearance      =   0  'Flat
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
         Left            =   2355
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1935
         Width           =   5100
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   720
         Left            =   2385
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4245
         Width           =   1350
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   720
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4245
         Width           =   1395
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   720
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4245
         Width           =   1425
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   720
         Left            =   5235
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4245
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   720
         Left            =   6735
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4245
         Width           =   1395
      End
      Begin VB.ComboBox cboCollege 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2355
         TabIndex        =   8
         Top             =   3375
         Width           =   7755
      End
      Begin VB.ComboBox cboagent 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2355
         TabIndex        =   0
         Top             =   495
         Width           =   5115
      End
      Begin VB.ComboBox cbosubject 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2355
         TabIndex        =   1
         Top             =   870
         Width           =   5115
      End
      Begin VB.ComboBox cboDist 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2355
         TabIndex        =   5
         Top             =   2280
         Width           =   5595
      End
      Begin VB.ComboBox cboCity 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2355
         TabIndex        =   6
         Top             =   2655
         Width           =   5595
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   3855
         Left            =   270
         TabIndex        =   13
         Top             =   5130
         Width           =   12945
         _cx             =   22834
         _cy             =   6800
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   7917545
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
      Begin MSMask.MaskEdBox txtDOB 
         Height          =   315
         Left            =   2355
         TabIndex        =   9
         Top             =   3795
         Width           =   1635
         _ExtentX        =   2900
         _ExtentY        =   572
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtAnniversaryDate 
         Height          =   315
         Left            =   5955
         TabIndex        =   10
         Top             =   3795
         Width           =   1635
         _ExtentX        =   2900
         _ExtentY        =   572
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   870
         Left            =   2340
         Top             =   4185
         Width           =   7305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 1 (Home):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   28
         Top             =   1590
         Width           =   1905
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Teacher Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   27
         Top             =   1230
         Width           =   1800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 2 :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   26
         Top             =   1965
         Width           =   1785
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   25
         Top             =   3795
         Width           =   1755
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Anniversary Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4095
         TabIndex        =   24
         Top             =   3810
         Width           =   1800
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "School Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   23
         Top             =   3420
         Width           =   1785
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "District          :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   22
         Top             =   2325
         Width           =   1815
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "City              :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   21
         Top             =   2685
         Width           =   1785
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   20
         Top             =   495
         Width           =   1815
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   19
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone              :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   18
         Top             =   3060
         Width           =   1770
      End
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   15885
      Top             =   225
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmTeacherDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Name1 As String
Dim city As String
Dim bb As Boolean
Dim collid As Integer
Dim rs1 As New ADODB.Recordset
Dim uniid As Integer
Private Sub cboagent_Click()
  fillGrid
End Sub
Private Sub cboagent_GotFocus()
  SendKeys "{f4}"
End Sub
Private Sub cboCity_GotFocus()
  SendKeys "{f4}"
End Sub
Private Sub cboCollege_GotFocus()
  SendKeys "{f4}"
End Sub
Private Sub cbodist_Click()
cboCollege.Clear

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT College.College, College.city, College.district " & _
"" & _
" FROM College where " & stringyear & " and College.district='" & Trim(cboDist) & "'  order by College.College", con, adOpenForwardOnly, adLockReadOnly
While rs1.EOF = False

If (rs1!District = rs1!city) Then
   cboCollege.AddItem rs1!college & "  " & rs1!city
Else
   cboCollege.AddItem rs1!college & "  " & rs1!city & " , " & rs1!District
End If

rs1.MoveNext
Wend

End Sub

Private Sub cboDist_GotFocus()
  SendKeys "{f4}"
End Sub
Private Sub cbosubject_Click()
  fillGrid
End Sub
Private Sub cbosubject_GotFocus()
SendKeys "{f4}"
End Sub
Private Sub cmdDelete_Click()

If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "delete from teacher where (" & stringyear & " and school='" & cboCollege.Text & "' and name='" & cboagent & "')"
   Call cmdRefresh_Click
End If
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdRefresh_Click()
'bb = False
Dim o As Object
    For Each o In Me
          If (TypeOf o Is textbox Or TypeOf o Is ComboBox) Then
          o.Text = ""
          End If
      Next
      
      cbosubject.Text = ""
      
      txtAnniversaryDate.Text = "__/__/____"
      txtDOB.Text = "__/__/____"
      
      fillGrid
      cboagent.SetFocus
End Sub
Private Sub cmdSave_Click()

saveData
Call cmdRefresh_Click

'End If
End Sub
Sub saveData()

On Error GoTo save1

If cboagent.Text = "" Then
   MsgBox "Please Select Agent !!", vbInformation
   cboagent.SetFocus
   Exit Sub
End If
        
If cbosubject.Text = "" Then
   MsgBox "Please Select subject !!", vbInformation
   cbosubject.SetFocus
   Exit Sub
End If
        
If cboCollege.Text = "" Then
   MsgBox "Please Select School !!", vbInformation
   cboCollege.SetFocus
   Exit Sub
End If
        
         
Set rs2 = New ADODB.Recordset
rs2.Open "select * from teacher where (" & stringyear & " and school='" & cboCollege.Text & "' and Name='" & txtName & "')", con, adOpenDynamic, adLockOptimistic
If rs2.EOF = True Then
   rs2.AddNew
End If

rs2!school = cboCollege.Text
rs2!Agent = cboagent.Text
rs2!Subject = cbosubject.Text
rs2!Name = txtName.Text
rs2!Address = txtAddress1.Text

If IsDate(txtDOB.Text) Then
   rs2!DOB = txtDOB.Text
End If

If IsDate(txtAnniversaryDate.Text) Then
   rs2!andate = txtAnniversaryDate.Text
End If

rs2!address2 = Trim(txtAddress2)
rs2!phone = txtphone.Text

rs2!city = Trim(cboCity)
rs2!District = Trim(cboDist)

rs2!fyear = session
rs2!setupid = setupid
rs2.update
          
          
Exit Sub

save1:

MsgBox "" & err.DESCRIPTION
     
End Sub
Sub search()
            
    Set rs1 = New Recordset
    rs1.Open "select * from teacherDetail where " & stringyear & " and auto=" & popupvalue4 & "", con
    If rs1.EOF = False Then
            
            txtCollageID.Text = rs1!collagelId
            txtCollageName.Text = rs1!college
            txtName.Text = rs1!Name
            txtAddress1.Text = rs1!address1
            txtAddress2.Text = rs1!address2
            txtAddress3.Text = rs1!address3
            txtcity.Text = rs1!city
            txtstate.Text = rs1!State
            txtphone.Text = rs1!phone
            txtFax.Text = rs1!FAX
            txtEmail.Text = rs1!email
            txtDOB.Text = rs1!DOB
            txtAnniversaryDate.Text = rs1!anniversaryDate
            txtCity1.Text = rs1!city1 & ""
            txtDist.Text = rs1!dist & ""
                        
    End If
     
            
            
End Sub


Private Sub Command1_Click()
    cr.Reset
    cr.ReportFileName = rptPath & "/teacherList.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    
    If (cboagent.Text <> "" And cbosubject.Text = "") Then
        cr.ReplaceSelectionFormula "{teacher.agent}='" & Me.cboagent.Text & "'"
    ElseIf (cboagent.Text <> "" And cbosubject.Text <> "") Then
        cr.ReplaceSelectionFormula "({teacher.agent}='" & Me.cboagent.Text & "' and {teacher.subject}='" & Me.cbosubject.Text & "')"
    ElseIf (cboagent.Text = "" And cbosubject.Text <> "") Then
        cr.ReplaceSelectionFormula "{teacher.subject}='" & Me.cbosubject.Text & "'"
    
    End If

    
    cr.WindowShowPrintBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowSearchBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
 
If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("txtCity1")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("txtDist")) Then
     SendKeys "{tab}"
End If

End If


End Sub
Sub fillGrid()

Dim rs_f As New ADODB.Recordset

If rs_f.State = 1 Then rs_f.close

If cboagent.Text = "" Then
   'rs_f.Open "select * from teacher order by school", CON, adOpenKeyset, adLockReadOnly
   Set rs_f = con.Execute("exec Search_Qry '4','" & session & "'," & main.setupid & "")
ElseIf (cboagent.Text <> "" And cbosubject.Text = "") Then
   rs_f.Open "select * from teacher where " & stringyear & " and agent='" & cboagent.Text & "' order by school", con, adOpenKeyset, adLockReadOnly
ElseIf (cboagent.Text <> "" And cbosubject.Text <> "") Then
   rs_f.Open "select * from teacher where (" & stringyear & " and agent='" & cboagent.Text & "' and subject='" & cbosubject & "') order by school", con, adOpenKeyset, adLockReadOnly

End If

Set vs.DataSource = rs_f


End Sub

Private Sub Form_Load()
bb = False

Screen.MousePointer = vbHourglass



If rs1.State = 1 Then rs1.close
rs1.Open "SELECT agentname" & _
" FROM AGENTMASTER where " & stringyear & " order by agentname", con, adOpenForwardOnly, adLockReadOnly
While rs1.EOF = False
    cboagent.AddItem rs1(0)
    rs1.MoveNext
Wend


If rs1.State = 1 Then rs1.close
rs1.Open "SELECT subject" & _
" FROM subjectname", con, adOpenForwardOnly, adLockReadOnly
While rs1.EOF = False
    cbosubject.AddItem Trim(rs1(0))
    rs1.MoveNext
Wend


'If rs1.State = 1 Then rs1.close
'rs1.Open "SELECT distinct(city)" & _
'" FROM city order by city", CON, adOpenStatic, adLockPessimistic
Set rs1 = con.Execute("exec Search_Qry '2','" & session & "'," & main.setupid & "")
While rs1.EOF = False
cboCity.AddItem Trim(rs1(0))
rs1.MoveNext
Wend

'If rs1.State = 1 Then rs1.close
'rs1.Open "SELECT distinct(District)" & _
'" FROM city order by District", CON, adOpenForwardOnly, adLockReadOnly

Set rs1 = con.Execute("exec Search_Qry '3','" & session & "'," & main.setupid & "")
While rs1.EOF = False
cboDist.AddItem Trim(rs1(0))
rs1.MoveNext
Wend

fillGrid
BackColorFrom Me

Screen.MousePointer = vbDefault

End Sub
Private Sub Form_Resize()

panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub txtAnniversaryDate_GotFocus()
txtAnniversaryDate.SelStart = 0
txtAnniversaryDate.SelLength = 10
End Sub

Private Sub txtCity1_GotFocus()
If PopUpValue1 <> "" Then
txtCity1.Text = PopUpValue1
txtName.SetFocus
PopUpValue1 = ""
End If
End Sub
Private Sub txtCollageName_GotFocus()
If PopUpValue1 <> "" Then
Name1 = ""
colid = ""
uniid = "0"
city = ""
txtCollageName.Text = PopUpValue1
txtCollageID.Text = popupvalue4
Name1 = PopUpValue1
collid = popupvalue4
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""
End If
End Sub

Private Sub txtCollageName_KeyDown(KeyCode As Integer, Shift As Integer)
Screen.MousePointer = vbHourglass

'If KeyCode = 13 Then
'SendKeys "{tab}"
'End If
''If KeyCode = 113 Then
''If bb = False Then
''    popuplistnew "select College,City,District,CollegeID  from College order by College", CON
''    bb = True
''Else
''   popuplist1.Visible = True
''End If
''End If
Screen.MousePointer = vbDefault
End Sub

Private Sub txtDist_GotFocus()
If PopUpValue1 <> "" Then
txtDist.Text = PopUpValue1
PopUpValue1 = ""
txtCity1.SetFocus
End If
End Sub

Private Sub txtDist_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 113 Or KeyCode = 13 Then
''   If txtCollageName.Text = "" Then
''     popuplist2 "select distinct(District) from College order by District", CON
''   Else
''      popuplist2 "select distinct(District) from College where College='" & txtCollageName.Text & "' order by District", CON
''   End If
''
''End If
End Sub

Private Sub txtDOB_GotFocus()
txtDOB.SelStart = 0
txtDOB.SelLength = 10
End Sub

Private Sub txtDOB_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'SendKeys "{tab}"
'End If
End Sub

Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'SendKeys "{tab}"
'End If
End Sub

Private Sub txtfax_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'SendKeys "{tab}"
'End If
End Sub

Private Sub txtName_GotFocus()
If PopUpValue1 > "" Then
txtName.Text = PopUpValue1
search
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""
popupvalue5 = ""
End If
End Sub
Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'SendKeys "{tab}"
'End If
'If KeyCode = 113 Then
'popuplist2 "select Name,College,City1 as City,auto from teacherdetail order by Name,College", CON
'End If
End Sub

Private Sub txtphone_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'SendKeys "{tab}"
'End If
End Sub

Private Sub txtstate_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'SendKeys "{tab}"
'End If
End Sub
Private Sub vs_DblClick()

cboCollege.Text = vs.TextMatrix(vs.RowSel, 1)
cboagent.Text = vs.TextMatrix(vs.RowSel, 4)
cbosubject.Text = vs.TextMatrix(vs.RowSel, 3)
txtName.Text = vs.TextMatrix(vs.RowSel, 2)
txtAddress1 = vs.TextMatrix(vs.RowSel, 5)
txtAddress2.Text = vs.TextMatrix(vs.RowSel, 6)
cboDist = vs.TextMatrix(vs.RowSel, 7)
cboCity = vs.TextMatrix(vs.RowSel, 8)
txtphone = vs.TextMatrix(vs.RowSel, 9)

If IsDate(vs.TextMatrix(vs.RowSel, 10)) Then
   txtDOB.Text = vs.TextMatrix(vs.RowSel, 10)
Else
   txtDOB.Text = "__/__/____"
End If

If IsDate(vs.TextMatrix(vs.RowSel, 11)) Then
   txtAnniversaryDate.Text = vs.TextMatrix(vs.RowSel, 11)
Else
   txtAnniversaryDate.Text = "__/__/____"
End If

End Sub
