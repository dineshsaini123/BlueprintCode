VERSION 5.00
Begin VB.Form frmBookSeller 
   Caption         =   "Book Seller"
   ClientHeight    =   5685
   ClientLeft      =   4275
   ClientTop       =   2520
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7890
   Begin VB.Frame Frame1 
      Height          =   5370
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   7740
      Begin VB.TextBox txtContact 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   285
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   5
         Top             =   3060
         Width           =   3525
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H80000008&
         Height          =   1080
         Left            =   90
         TabIndex        =   18
         Top             =   4200
         Width           =   7575
         Begin VB.CommandButton cmdAdd_1 
            BackColor       =   &H00FFFFC0&
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
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   135
            Width           =   1230
         End
         Begin VB.CommandButton cmdSave_2 
            BackColor       =   &H00FFFFC0&
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
            Left            =   1290
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   135
            Width           =   1230
         End
         Begin VB.CommandButton cmdDelete_3 
            BackColor       =   &H00FFFFC0&
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
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   135
            Width           =   1230
         End
         Begin VB.CommandButton cmdEdit_4 
            BackColor       =   &H00FFFFC0&
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
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   135
            Width           =   1230
         End
         Begin VB.CommandButton cmdExit_12 
            BackColor       =   &H00FFFFC0&
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
            Left            =   6255
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   135
            Width           =   1230
         End
         Begin VB.CommandButton cmdSearch 
            BackColor       =   &H00FFFFC0&
            Caption         =   "S&earch"
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
            Left            =   4995
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   135
            Width           =   1230
         End
      End
      Begin VB.TextBox txtRepresentative 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Pub_code"
         Height          =   285
         Left            =   1755
         TabIndex        =   0
         Top             =   600
         Width           =   3525
      End
      Begin VB.TextBox txtRId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Pub_code"
         Height          =   285
         Left            =   5295
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   585
         Width           =   1110
      End
      Begin VB.TextBox txtAddress1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   285
         Left            =   1755
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1125
         Width           =   3570
      End
      Begin VB.TextBox txtAddress2 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   285
         Left            =   1755
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1545
         Width           =   3570
      End
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Pub_code"
         Height          =   285
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1965
         Width           =   3540
      End
      Begin VB.TextBox txtCityId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Pub_code"
         Height          =   285
         Left            =   5310
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1965
         Width           =   1110
      End
      Begin VB.TextBox txtpin 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   285
         Left            =   3960
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2280
         Width           =   1320
      End
      Begin VB.TextBox txtPhone1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   285
         Left            =   1755
         MaxLength       =   30
         TabIndex        =   6
         Top             =   3450
         Width           =   3525
      End
      Begin VB.TextBox txtDist 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Pub_code"
         Height          =   285
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2280
         Width           =   1785
      End
      Begin VB.TextBox txtState 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Pub_code"
         Height          =   285
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2595
         Width           =   3540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   6435
         Picture         =   "frmBookSeller.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Add City"
         Top             =   1920
         Width           =   510
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   285
         Left            =   1755
         MaxLength       =   30
         TabIndex        =   7
         Top             =   3780
         Width           =   3525
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   405
         Index           =   4
         Left            =   780
         TabIndex        =   30
         Top             =   2940
         Width           =   885
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Representative :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   2
         Left            =   315
         TabIndex        =   29
         Top             =   585
         Width           =   1470
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address -1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   765
         TabIndex        =   28
         Top             =   1125
         Width           =   1020
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "City/Town :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   12
         Left            =   765
         TabIndex        =   27
         Top             =   1965
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "District :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   13
         Left            =   765
         TabIndex        =   26
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sate :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   14
         Left            =   765
         TabIndex        =   25
         Top             =   2595
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pin :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   15
         Left            =   3555
         TabIndex        =   24
         Top             =   2280
         Width           =   510
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   17
         Left            =   765
         TabIndex        =   23
         Top             =   3450
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address -2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   1
         Left            =   780
         TabIndex        =   22
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   3
         Left            =   780
         TabIndex        =   21
         Top             =   3780
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmBookSeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean
Private Sub cmdAdd_1_Click()

txtRId = MaxSNo("BookSeler", "BookSelerID", "BookSeler")

'txtRepresentative.SetFocus

txtRepresentative = ""
txtNikeName = ""
txtAddress1 = ""
txtAddress2 = ""
txtCityId = ""
txtPhone1 = ""
txtPhone2 = ""
txtpin = ""
txtCity = ""
txtDist = ""
txtCityId = ""
txtState = ""
txtContact = ""
'cboDesig.ListIndex = -1
txtEmail = ""



edit1 = False

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

   
   
End Sub

Private Sub cmdDelete_3_Click()
CON.BeginTrans
CON.Execute "delete from  [BookSeler] where BookSelerID='" & txtRId & "'"
CON.CommitTrans
cmdAdd_1_Click
txtRepresentative.SetFocus
End Sub

Private Sub cmdEdit_4_Click()
edit1 = True
cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True
cmdSave_2.SetFocus

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdSave_2_Click()


If txtRepresentative = "" Then
   MsgBox "Enter Name. ...", vbInformation
   txtRepresentative.SetFocus
   Exit Sub
End If
   
'If txtNikeName = "" Then
'   MsgBox "Enter Code Name. ...", vbInformation
'   txtNikeName.SetFocus
'   Exit Sub
'End If


If txtCityId = "" Then
   MsgBox "Select City Name ...", vbInformation
   txtCityId.SetFocus
   Exit Sub
End If


CON.BeginTrans

If edit1 = True Then
   CON.Execute "delete from  [BookSeler] where BookSelerID='" & txtRId & "'"
End If

   

CON.Execute "INSERT INTO  [BookSeler]" & _
           "([BookSelerID]" & _
           ",[BookSeler]" & _
           ",[Add1]" & _
           ",[Add2]" & _
           ",[CityID]" & _
           ",[Phone],[Pin],[Email],[ContectPerson])" & _
     "Values" & _
           "('" & txtRId & "'" & _
           ",'" & Trim(txtRepresentative) & "'" & _
           "" & _
           ",'" & txtAddress1 & "'" & _
           ",'" & txtAddress2 & "'" & _
           ",'" & txtCityId & "'" & _
           ",'" & txtPhone1 & "','" & txtpin & "','" & txtEmail.Text & "','" & Trim(txtContact) & "')"
CON.CommitTrans

MsgBox "Date Saved ....", vbInformation
cmdSave_2.Enabled = False

Call cmdAdd_1_Click
txtRepresentative.SetFocus


End Sub

Private Sub cmdSearch_Click()

   tblNo = 50
   frmSearchItem.Show
   cmdSave_2.Enabled = False
   cmdEdit_4.Enabled = True

End Sub
Private Sub cmdSearch_GotFocus()
  
  If rs.State = 1 Then rs.Close
  rs.Open "select * from  [BookSeler] where BookSelerID='" & PopUpValue1 & "'", CON
  If rs.EOF = False Then
  
  
   'cmdSave_2.Enabled = False
   'cmdEdit_4.Enabled = True
   'cmdDelete_3.Enabled = True

  
     
   txtRId = rs![BookSelerID]
   txtRepresentative = rs![BookSeler]
   txtAddress1 = rs![ADD1] & ""
   txtAddress2 = rs![ADD2] & ""
   txtpin = rs!pin
   txtCityId = rs![CityID] & ""
   txtPhone1 = rs![Phone] & ""
   txtEmail = rs![email] & ""
   txtContact = rs![ContectPerson] & ""
''   If Not IsNull(rs!Designation) Then
''   cboDesig.Text = rs!Designation
''   End If
   
   
   rs_map.MoveFirst
   rs_map.Find "[CityID]='" & rs![CityID] & "'"
   If rs_map.EOF = False Then
      txtCity = rs_map!city & ""
      txtDist = rs_map!district & ""
      txtState = rs_map!State & ""
   End If
   
   
  PopUpValue1 = ""
  txtRepresentative.SetFocus
   
   
  End If

End Sub
Private Sub Command1_Click()
   tblNo = 6
   frmSearchItem.Show
End Sub

Private Sub Command1_GotFocus()
If PopUpValue1 <> "" Then
   
    txtCity = PopUpValue2
    txtCityId = PopUpValue1
    
    If rs.State = 1 Then rs.Close
    rs.Open "select [District],[State] FROM  [CityView] " & _
    "where [CityID]='" & PopUpValue1 & "'", CON
    If rs.EOF = False Then
       txtDist = rs(0)
       txtState = rs(1)
    End If
    
    txtpin.SetFocus
End If

PopUpValue1 = ""
PopUpValue2 = ""

End Sub

Private Sub Form_Activate()
'cmdAdd_1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub Form_Load()

'Me.Top = 700
'Me.Left = 500


'fillcombo cboDesig, "Designation", "Designation", CON
cmdAdd_1_Click
End Sub

Private Sub txtcity_GotFocus()
HIT
If PopUpValue1 <> "" Then
   
    txtCity = PopUpValue2
    txtCityId = PopUpValue1
    
    If rs.State = 1 Then rs.Close
    rs.Open "select [District],[State] FROM  [CityView] " & _
    "where [CityID]='" & PopUpValue1 & "'", CON
    If rs.EOF = False Then
       txtDist = rs(0)
       txtState = rs(1)
    End If
    
    txtpin.SetFocus
End If

PopUpValue1 = ""
PopUpValue2 = ""

End Sub

Private Sub txtcity_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then Exit Sub
   If KeyCode = 13 Then Exit Sub
   
   tblNo = 6
   frmSearchItem.Show
   
End Sub




