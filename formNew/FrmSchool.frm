VERSION 5.00
Begin VB.Form FrmSchool 
   Caption         =   "School Master"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame panel 
      Caption         =   "School Master"
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
      Height          =   7380
      Left            =   135
      TabIndex        =   16
      Top             =   135
      Width           =   11085
      Begin VB.CommandButton cmdpr 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Previous"
         Height          =   465
         Left            =   6435
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4950
         Width           =   1020
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next"
         Height          =   465
         Left            =   7515
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4950
         Width           =   1020
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   270
         ScaleHeight     =   825
         ScaleWidth      =   8310
         TabIndex        =   34
         Top             =   5760
         Width           =   8310
         Begin VB.CommandButton close 
            BackColor       =   &H00FFFFFF&
            Cancel          =   -1  'True
            Caption         =   "&Exit"
            Height          =   675
            Left            =   7155
            Picture         =   "FrmSchool.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   75
            Width           =   1095
         End
         Begin VB.CommandButton REPORTCD 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   675
            Left            =   5970
            Picture         =   "FrmSchool.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   75
            Width           =   1095
         End
         Begin VB.CommandButton Abandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Abandon"
            Height          =   675
            Left            =   4788
            Picture         =   "FrmSchool.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   75
            Width           =   1095
         End
         Begin VB.CommandButton Del 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            Height          =   675
            Left            =   3606
            Picture         =   "FrmSchool.frx":1D52
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   75
            Width           =   1095
         End
         Begin VB.CommandButton save 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   675
            Left            =   1242
            Picture         =   "FrmSchool.frx":2936
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   75
            Width           =   1095
         End
         Begin VB.CommandButton Help 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   675
            Left            =   60
            Picture         =   "FrmSchool.frx":351A
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   75
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Modify"
            Enabled         =   0   'False
            Height          =   675
            Left            =   2424
            Picture         =   "FrmSchool.frx":40FE
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   75
            Width           =   1095
         End
      End
      Begin VB.ComboBox cboBookCat 
         Height          =   315
         Left            =   1470
         TabIndex        =   5
         Top             =   2745
         Width           =   3585
      End
      Begin VB.TextBox txtcollegeUpdate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   31
         Top             =   5085
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add Dist."
         Height          =   345
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1980
         Width           =   780
      End
      Begin VB.TextBox txtPhone1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1470
         MaxLength       =   25
         TabIndex        =   7
         Top             =   3540
         Width           =   3585
      End
      Begin VB.TextBox txtadd1 
         Appearance      =   0  'Flat
         Height          =   750
         Left            =   1470
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   1155
         Width           =   3585
      End
      Begin VB.TextBox txtCollege 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   1
         Top             =   780
         Width           =   3645
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1995
         Width           =   3600
      End
      Begin VB.TextBox txtPin 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3150
         Width           =   3585
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   8
         Top             =   3915
         Width           =   3600
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1470
         MaxLength       =   25
         TabIndex        =   10
         Top             =   4695
         Width           =   3585
      End
      Begin VB.TextBox txtWebSite 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1470
         MaxLength       =   35
         TabIndex        =   9
         Top             =   4305
         Width           =   3600
      End
      Begin VB.TextBox txtPname 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7110
         MaxLength       =   50
         TabIndex        =   11
         Top             =   660
         Width           =   2370
      End
      Begin VB.TextBox city 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1470
         MaxLength       =   100
         TabIndex        =   4
         Top             =   2385
         Width           =   3600
      End
      Begin VB.TextBox collegeid 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1470
         TabIndex        =   0
         Top             =   405
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.TextBox prphone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7110
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1080
         Width           =   2385
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   915
         Left            =   270
         Top             =   5715
         Width           =   8385
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F2 For search the record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   135
         TabIndex        =   33
         Top             =   6975
         Width           =   3525
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   32
         Top             =   2685
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   " Phone Nos. : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   3525
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   " Address : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   27
         Top             =   1125
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   " College :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   765
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "District"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   " PinCode:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   24
         Top             =   3090
         Width           =   810
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "FAX :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   23
         Top             =   3900
         Width           =   480
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   22
         Top             =   4710
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "School URL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   21
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Principle Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5610
         TabIndex        =   20
         Top             =   660
         Width           =   1410
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label Label7 
         Caption         =   "College ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   165
         TabIndex        =   18
         Top             =   405
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Nos."
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
         Left            =   5640
         TabIndex        =   17
         Top             =   1095
         Width           =   1395
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   11700
      Top             =   3375
   End
   Begin VB.TextBox txtCity 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   11250
      TabIndex        =   13
      Top             =   135
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "City Id :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9675
      TabIndex        =   15
      Top             =   150
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label LABELUSERNAME 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   7875
      Width           =   7785
   End
End
Attribute VB_Name = "FrmSchool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim addmode As Boolean
Dim labeltrue As Boolean
Dim update As Boolean
Dim I As Integer
Dim chkcityid As Boolean
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim kk1 As Integer


Private Sub ABANDON_Click()
clearfield
disable
Me.txtCollege.SetFocus
End Sub
Private Sub city_GotFocus()
    
Label28.Visible = True
labeltrue = True

    If PopUpValue1 <> "" Then
    Me.city = PopUpValue1
       End If
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
End Sub


Private Sub city_LostFocus()
Label28.Visible = False
labeltrue = False
End Sub

Private Sub close_Click()
 Unload Me
End Sub
Private Sub cmdAdd_Click()
         'popuplistModel10 "select District from College where " & stringyear & " group by District", CON
         popuplist_client "select District from College where " & stringyear & " group by District", CCON
End Sub
Private Sub cmdAdd_GotFocus()
   If PopUpValue1 <> "" Then
      txtArea = PopUpValue1
      PopUpValue1 = ""
      cmdpr.Enabled = False
      save.Enabled = True
      update = True
      enable
      Command1.Enabled = True
      
      Set rss = New ADODB.Recordset
      rss.Open "select * from college where " & stringyear & " and district='" + txtArea + "' order by College", CON, adOpenDynamic, adLockPessimistic

      
   End If
   
End Sub
Sub SearchData()
On Error GoTo aa:


If kk1 = 1 Then
   rss.MoveNext
Else
   rss.MovePrevious
End If
        
Me.collegeid.Text = rss("collegeid")
Me.txtCollege.Text = rss("College") & ""
Me.txtadd1.Text = rss("Add1") & ""
Me.txtArea.Text = rss("district") & ""
Me.city.Text = IIf(IsNull(rss("city")), "", rss!city)
Me.txtPin.Text = rss("Pin") & ""
Me.txtPhone1.Text = rss("Phone1") & ""
Me.txtWebSite.Text = IIf(IsNull(rss("WebSite")), "", rss!website)
Me.txtFax.Text = rss("Fax") & ""
Me.txtEmail.Text = rss("Email") & ""
Me.prphone.Text = IIf(IsNull(rss("pphone")), "", rss!pphone)
Me.txtPname.Text = IIf(IsNull(rss("Principalname")), "", rss!principalname)

Exit Sub
aa:
MsgBox "" & Err.DESCRIPTION

End Sub
Private Sub cmdNext_Click()


kk1 = 1
cmdpr.Enabled = True
SearchData


End Sub

Private Sub cmdpr_Click()
kk1 = 2
SearchData
End Sub

Private Sub Command1_Click()
Dim RS As New ADODB.Recordset




If Me.txtCollege.Text = "" Then
MsgBox "Please fill the college name"
Me.txtCollege.SetFocus
Exit Sub
End If

If Me.txtArea = "" Then
MsgBox "Please fill the college city name"
Me.txtArea.SetFocus
Exit Sub
End If


If Me.city = "" Then
MsgBox "Please fill the college city name"
Me.city.SetFocus
Exit Sub
End If

Set RS = New ADODB.Recordset
  Dim I As Integer


If RS.State = 1 Then RS.close
            RS.Open "select * from districts where " & stringyear, CON, adOpenDynamic, adLockOptimistic
                RS.Find "districtname='" + Trim(UCase(Me.txtArea.Text)) + "'"
                If Not RS.EOF Then
                Else
                    RS.AddNew
                    RS(0) = Trim(UCase(Me.txtArea))
                    RS.update
                End If
        
        
        Dim cit As New ADODB.Recordset
        
        
              If RS.State = 1 Then RS.close
              RS.Open "select * from city where " & stringyear, CON, adOpenDynamic, adLockOptimistic
              cit.Open "select * from city where " & stringyear & " and district='" + Trim(UCase(Me.txtArea.Text)) + "'", CON, adOpenStatic, adLockReadOnly
              'cit.Find "district='" + Trim(UCase(Me.txtArea.Text)) + "'"
              If cit.EOF = False Then
                   
                   RS.Find "city='" + Trim(UCase(Me.city.Text)) + "'"
                   If RS.EOF = True Then
                    RS.AddNew
                    RS(1) = Trim(UCase(Me.city))
                    RS(2) = Trim(UCase(Me.txtArea.Text))
                    RS.update
                    End If
                Else
                    RS.AddNew
                    RS(1) = Trim(UCase(Me.city))
                    RS(2) = Trim(UCase(Me.txtArea.Text))
                    RS.update
            End If
            






If update = True Then

If RS.State = 1 Then RS.close
RS.Open "select * from College where " & stringyear & " and collegeid= " & collegeid & "", CON, adOpenDynamic, adLockOptimistic
If rs1.State = 1 Then rs1.close
rs1.Open "select cityid from city where " & stringyear & " and city='" & city.Text & "' and district='" & txtArea.Text & "'", CON, adOpenStatic, adLockReadOnly
 
Else
If RS.State = 1 Then RS.close
RS.Open "select * from College where " & stringyear & "", CON, adOpenDynamic, adLockOptimistic
If rs1.State = 1 Then rs1.close
rs1.Open "select cityid from city where " & stringyear & " and city='" & city.Text & "' and district='" & txtArea.Text & "'", CON, adOpenStatic, adLockReadOnly
If addmode = True Then
    RS.AddNew
    End If
End If
   
   If RS.RecordCount >= 0 Then
         
            
         If rs1.RecordCount > 0 Then
         RS("cityid") = rs1("cityid")
         End If
               
         RS("College") = txtCollege & ""
         RS("Add1") = txtadd1 & ""
         RS("district") = txtArea & ""
         RS("city") = city & ""
         RS("Pin") = txtPin & ""
         RS("Phone1") = txtPhone1 & ""
         RS("WebSite") = txtWebSite & ""
          RS("Fax") = txtFax & ""
         RS("Email") = txtEmail & ""
         RS("pphone") = prphone & ""
         RS("Principalname") = txtPname & ""
         RS.update
         
         RS.close
        

End If

End Sub

Private Sub Del_Click()



Set RS = New ADODB.Recordset
RS.Open "select * from college where " & stringyear & " and collegeid= " & CInt(Me.collegeid.Text) & "", CON, adOpenDynamic, adLockPessimistic
If RS.RecordCount > 0 Then
If Me.txtCollege.Text <> "" Then
CON.Execute "delete * from college where " & stringyear & " and collegeid=" & CInt(Me.collegeid.Text) & ""
End If
Else
MsgBox "Please select a proper record"
End If
clearfield
Me.txtCollege.SetFocus


End Sub

Private Sub Form_Activate()
'Help.SetFocus
Me.txtCollege.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub
Private Sub Form_Load()
On Error Resume Next

disable
save.Visible = True

'popuplist11 "Select college,cityid,pin,collegeid from college order by college ", CON, , True

If RS.State = 1 Then RS.close
RS.Open "select * from bookcategory order by bookcategory", CON
While RS.EOF = False
cboBookCat.AddItem RS(0)
RS.MoveNext
Wend

BackColorFrom Me


End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub Help_Click()
clearfield
enable
 'CmdUpdate.Visible = False
save.Enabled = True
Me.txtCollege.SetFocus
addmode = True
update = False
End Sub

Private Sub save_Click()

Dim RS As New ADODB.Recordset

If cboBookCat.Text = "" Then
MsgBox "Enter Category ...", vbCritical
cboBookCat.SetFocus

Exit Sub
End If



If Me.txtCollege.Text = "" Then
MsgBox "Please fill the college name"
Me.txtCollege.SetFocus
Exit Sub
End If

If Me.txtArea = "" Then
MsgBox "Please fill the college city name"
Me.txtArea.SetFocus
Exit Sub
End If


If Me.city = "" Then
MsgBox "Please fill the college city name"
Me.city.SetFocus
Exit Sub
End If

Set RS = New ADODB.Recordset
  Dim I As Integer


If RS.State = 1 Then RS.close
        
        RS.Open "select * from districts where " & stringyear & " and districtname='" + Trim(UCase(Me.txtArea.Text)) + "'", CON, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
        Else
            RS.AddNew
            RS(0) = Trim(UCase(Me.txtArea))
            RS.update
        End If
        
        
        Dim cit As New ADODB.Recordset
        
        
         'If RS.State = 1 Then RS.close
         'RS.Open "select * from city where " & stringyear, CON, adOpenDynamic, adLockOptimistic
              
              'If cit.State = 1 Then cit.close
              'cit.Open "select * from city where " & stringyear & " and district='" + Trim(UCase(Me.txtArea.Text)) + "'", CON, adOpenStatic, adLockReadOnly
              'cit.Find "district='" + Trim(UCase(Me.txtArea.Text)) + "'"
               
               If RS.State = 1 Then RS.close
               RS.Open "select * from city where " & stringyear & " and city='" + Trim(UCase(Me.city.Text)) + "' and district='" + Trim(UCase(Me.txtArea.Text)) + "'", CON, adOpenDynamic, adLockOptimistic

              
              If RS.EOF = False Then
                   
                   'RS.Find "city='" + Trim(UCase(Me.city.Text)) + "'"
                   'If RS.EOF = True Then
                    'RS.AddNew
                    RS(1) = Trim(UCase(Me.city))
                    RS(2) = Trim(UCase(Me.txtArea.Text))
                    RS!fyear = session
                    RS!setupid = fyear
                    RS.update
                    
                    'End If
                Else
                    RS.AddNew
                    RS(1) = Trim(UCase(Me.city))
                    RS(2) = Trim(UCase(Me.txtArea.Text))
                    RS!fyear = session
                    RS!setupid = fyear

                    RS.update
            End If
            













If update = True Then

If RS.State = 1 Then RS.close
RS.Open "select * from College where " & stringyear & " and collegeid= " & collegeid & "", CON, adOpenDynamic, adLockOptimistic
If rs1.State = 1 Then rs1.close
rs1.Open "select cityid from city where " & stringyear & " and city='" & city.Text & "' and district='" & txtArea.Text & "'", CON, adOpenStatic, adLockReadOnly
         
Else
If RS.State = 1 Then RS.close
RS.Open "select * from College where " & stringyear & "", CON, adOpenDynamic, adLockOptimistic
If rs1.State = 1 Then rs1.close
rs1.Open "select cityid from city where " & stringyear & " and city='" & city.Text & "' and district='" & txtArea.Text & "'", CON, adOpenStatic, adLockReadOnly
If addmode = True Then
    RS.AddNew
    End If
End If
   
   If RS.RecordCount >= 0 Then
         
            
         If rs1.RecordCount > 0 Then
         RS("cityid") = rs1("cityid")
         End If
               
         RS("College") = txtCollege & ""
         RS("Add1") = txtadd1 & ""
         RS("district") = txtArea & ""
         RS("city") = city & ""
         RS("Pin") = txtPin & ""
         RS("Phone1") = txtPhone1 & ""
         RS("WebSite") = txtWebSite & ""
         RS("Fax") = txtFax & ""
         RS("Email") = txtEmail & ""
         RS("pphone") = prphone & ""
         RS("Principalname") = txtPname & ""
         
         
         RS("states") = Trim(cboBookCat.Text)
         RS.update
        
        RS.close
        
        
        
        CON.Execute "update info set scname='" & Trim(Me.txtCollege) & "'" & _
        " where " & stringyear & " and city= '" & Trim(UCase(city)) & "' and district='" & txtArea.Text & "' and scname='" & txtcollegeUpdate.Text & "'"
        
        If collegeid <> "" Then
        CON.Execute "update teacherDetail set College='" & Trim(Me.txtCollege) & "' where " & stringyear & " and collagelId= " & collegeid & ""
        End If
        
        
        MsgBox "The Record Saved Successfully", vbInformation
        
clearfield
Me.txtCollege.SetFocus
save.Enabled = False
addmode = False
collegeid = 0
update = False
disable
End If
End Sub
Private Sub clearfield()
 For Each ctrl In Me.Controls
 If TypeOf ctrl Is textbox Then
   ctrl.Text = ""
 End If
 Next
 
 cboBookCat.Text = ""
 
' chkMarked.value = False
End Sub

Private Sub Timer1_Timer()
If labeltrue = True Then
Label28.Visible = Not Label28.Visible
Label28.ForeColor = RGB(255, 255, 255) * Rnd()
End If
End Sub

Private Sub txtadd1_KeyPress(KeyAscii As Integer)

''''If txtadd1.Text = "" Then
''''If KeyAscii = 13 Then
''''KeyAscii = 0
''''End If
''''End If

End Sub

Private Sub txtArea_GotFocus()
Label28.Visible = True
labeltrue = True
    If PopUpValue1 <> "" Then
    Me.txtArea = PopUpValue1
    End If
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
End Sub

Private Sub txtArea_KeyDown(KeyCode As Integer, Shift As Integer)
'''If KeyCode = 113 Then
'''popuplist12 "Select DISTRICTNAME from districts order by DISTRICTNAME ", CON
'''End If
End Sub

Private Sub txtArea_LostFocus()
Label28.Visible = False
labeltrue = False
End Sub
Private Sub txtCollege_GotFocus()
Dim colname As String
Label28.Visible = True
labeltrue = True

    If PopUpValue2 <> "" Then
      
      'colname = popuplist1.itemname
      
      
      If colname <> "" Then
      popuplist_client "Select city,college,district,collegeid from college where " & stringyear & "", CCON
      
      End If
         
         If PopUpValue1 <> "" Then
                lstfocus
                collegeid = CInt(popupvalue4)
                update = True
                enable
                colname = ""
           End If
    
    save.Enabled = True
    End If
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""

'popuplist1.itemname = ""
'popuplist1.ListView1.SelectedItem.Selected = False
End Sub

Private Sub txtCollege_KeyDown(KeyCode As Integer, Shift As Integer)
Dim colname As String
  
  If KeyCode = 113 Then
 
   popuplist_client "Select college,city,district,collegeid from  college where " & stringyear & " order by college ", CCON
   
  End If
End Sub

Private Sub txtCollege_LostFocus()
Label28.Visible = False
labeltrue = False
End Sub

Private Sub lstfocus()

Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.close
RS.Open "select * from college where " & stringyear & " and collegeid=" + popupvalue4 + "", CON, adOpenDynamic, adLockPessimistic
If RS.RecordCount > 0 Then
        'Me.txtUid.Text = RS("UniversityID")
        Me.collegeid.Text = RS("collegeid")
        Me.txtCollege.Text = RS("College")
        Me.txtcollegeUpdate = RS("College")
        Me.txtadd1.Text = RS("Add1")
        'Me.txtAdd2.Text = RS("Add2")
        Me.txtArea.Text = RS("district")
        Me.city.Text = IIf(IsNull(RS("city")), "", RS!city)
        Me.txtPin.Text = RS("Pin")
        Me.txtPhone1.Text = RS("Phone1")
        'Me.txtPhone2.Text = RS("Phone2")
        Me.txtWebSite.Text = IIf(IsNull(RS("WebSite")), "", RS!website)
        'Me.txtLand.Text = RS("Landmark")
        'Me.txtRating.Text = RS("Rating")
        Me.txtFax.Text = RS("Fax")
        Me.txtEmail.Text = RS("Email")
        Me.prphone.Text = IIf(IsNull(RS("pphone")), "", RS!pphone)
                'If RS("IsMarked") = True Then
                'Me.chkMarked.value = True
                'Else
                'Me.chkMarked.value = False
                'End If
        'Me.txtSummer1.Text = RS("TimingSummerI")
        'Me.txtSummer2.Text = RS("TimingSummerII")
        'Me.txtWinter1.Text = RS("TimingWinterI")
        'Me.txtWinter2.Text = RS("TimingWinterII")
        'Me.txtWeeklyh.Text = RS("WeeklyHoliday")
        'Me.txtVsummer.Text = RS("VacationSummer")
        'Me.txtVWinter.Text = RS("VacationWinter")
        'Me.txtBookf.Text = RS("BookFinalBy")
        'Me.txtBookfm.Text = RS("BookFinalMonth")
        'Me.txtConvensing.Text = RS("ConvensingMonth")
        Me.txtPname.Text = IIf(IsNull(RS("Principalname")), "", RS!principalname)
        cboBookCat.Text = RS("states") & ""
        
        'Me.txtSid.Text = RS("BookSellerID")
     
End If
End Sub
Sub disable()
Me.collegeid.Enabled = False
Me.txtadd1.Enabled = False
Me.txtArea.Enabled = False
Me.city.Enabled = False
Me.txtPhone1.Enabled = False
Me.txtFax.Enabled = False
Me.txtEmail.Enabled = False
Me.txtPin.Enabled = False
Me.txtWebSite.Enabled = False
Me.txtPname.Enabled = False
Me.prphone.Enabled = False
End Sub
Sub enable()
Me.collegeid.Enabled = True
Me.txtadd1.Enabled = True
Me.txtArea.Enabled = True
Me.city.Enabled = True
Me.txtPhone1.Enabled = True
Me.txtFax.Enabled = True
Me.txtEmail.Enabled = True
Me.txtPin.Enabled = True
Me.txtWebSite.Enabled = True
Me.txtPname.Enabled = True
Me.prphone.Enabled = True
End Sub
