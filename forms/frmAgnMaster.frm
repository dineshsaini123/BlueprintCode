VERSION 5.00
Begin VB.Form frmAgnMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agent Master"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtadd 
      Height          =   315
      Left            =   840
      MaxLength       =   50
      TabIndex        =   1
      Top             =   780
      Width           =   3720
   End
   Begin VB.TextBox txtName1 
      Height          =   315
      Left            =   840
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ListBox ListDis2 
      Height          =   2010
      ItemData        =   "frmAgnMaster.frx":0000
      Left            =   7380
      List            =   "frmAgnMaster.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   300
      Width           =   2235
   End
   Begin VB.ListBox ListDis1 
      Height          =   2010
      ItemData        =   "frmAgnMaster.frx":0004
      Left            =   4680
      List            =   "frmAgnMaster.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   315
      Width           =   2190
   End
   Begin VB.CommandButton Removecd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6960
      Picture         =   "frmAgnMaster.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   885
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton Addcd 
      Height          =   540
      Left            =   6960
      Picture         =   "frmAgnMaster.frx":034A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   300
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   3720
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   270
      TabIndex        =   11
      Top             =   3060
      Width           =   7755
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
         Height          =   720
         Left            =   3630
         Picture         =   "frmAgnMaster.frx":068C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1170
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
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
         Height          =   720
         Left            =   5115
         Picture         =   "frmAgnMaster.frx":1270
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   1245
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
         Height          =   720
         Left            =   6390
         Picture         =   "frmAgnMaster.frx":1E54
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   90
         Width           =   1245
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
         Height          =   720
         Left            =   3855
         Picture         =   "frmAgnMaster.frx":2A38
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   1245
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
         Height          =   720
         Left            =   2580
         Picture         =   "frmAgnMaster.frx":2E45
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   90
         Width           =   1245
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
         Height          =   720
         Left            =   1320
         Picture         =   "frmAgnMaster.frx":3A29
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   1245
      End
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
         Height          =   720
         Left            =   45
         Picture         =   "frmAgnMaster.frx":460D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   1245
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   1050
      Left            =   225
      Top             =   3015
      Width           =   7890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   420
      Width           =   510
   End
End
Attribute VB_Name = "frmAgnMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean

Private Sub Addcd_Click()
If ListDis1.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To ListDis1.ListCount - 1
        If ListDis1.Selected(I) Then
                ListDis2.AddItem ListDis1.List(I)
                delitem = I
         End If
    Next
    ListDis1.RemoveItem delitem
End If
End Sub

Private Sub cmdAdd_1_Click()

'txtRId = MaxSNo("Rep", "RepId", "Rep")
'txtRepresentative.SetFocus
   
'txtRId = ""

txtname = ""
txtadd1 = ""
txtAdd2 = ""

txtAdd = ""


edit1 = False

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

txtname.SetFocus

   
   
End Sub

Private Sub cmdDelete_3_Click()
CON.BeginTrans
CON.Execute "delete from  agentmaster where agentname='" & txtname & "' and " & stringyear
CON.CommitTrans
cmdAdd_1_Click
End Sub

Private Sub cmdEdit_4_Click()

cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True
cmdSave_2.SetFocus

edit1 = True

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub
Private Sub cmdSave_2_Click()


If txtname = "" Then
   MsgBox "Enter Agent Name. ...", vbInformation
   txtname.SetFocus
   Exit Sub
End If
   





If edit1 = True Then
   CON.Execute "update DISTRICTS set AGENTNAME='" & Trim(txtname) & "' where (AGENTNAME='" & txtName1 & "' and " & stringyear & ")"
   CON.Execute "update AGENTMASTER set address='" & Trim(txtAdd) & "' where (AGENTNAME='" & txtName1 & "' and " & stringyear & ")"
Else

CON.BeginTrans
CON.Execute "INSERT INTO  [AGENTMASTER]" & _
           "([AGENTNAME]" & _
           ",[address]" & _
           ",[Fyear]" & _
           ",[setupid]" & _
     ") Values" & _
           "('" & txtname & "'" & _
           ",'" & Trim(txtAdd) & "'" & _
           ",'" & main.session & "'" & _
           ",'" & main.setupid & "')"
        
CON.CommitTrans


End If

   



CON.Execute "update DISTRICTS set AGENTNAME='' where (AGENTNAME='" & txtname & "' and " & stringyear & ")"
For I = 0 To ListDis2.ListCount - 1
   CON.Execute "update DISTRICTS set AGENTNAME='" & Trim(txtname) & "' where  (DISTRICTNAME='" & ListDis2.List(I) & "' and " & stringyear & ")"
Next



MsgBox "Date Saved ....", vbInformation
cmdSave_2.Enabled = False

'Call cmdAdd_1_Click
   


End Sub

Private Sub cmdSearch_Click()

   popuplistModel10 "select distinct Agentname,Address from [AGENTMASTER] where  " & stringyear & " order by Agentname", CON
 
   cmdSave_2.Enabled = False
   cmdEdit_4.Enabled = True

End Sub
Private Sub cmdSearch_GotFocus()
  
  
  If PopUpValue1 <> "" Then
  
  txtname = PopUpValue1
  txtName1 = PopUpValue1
  
  txtAdd = PopUpValue2
  'txtAdd2 = PopUpValue3
  
  addDis2
  
  End If
   
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  
  cmdEdit_4.Enabled = True
  cmdSave_2.Enabled = False
  

End Sub
Private Sub Command1_Click()
   tblNo = 6
   frmSearchItem.Show
End Sub

Private Sub Command1_GotFocus()
If PopUpValue1 <> "" Then
   
    txtCity = PopUpValue2
    txtCityID = PopUpValue1
    
    If RS.State = 1 Then RS.close
    RS.Open "select [District],[State] FROM  [CityView] " & _
    "where " & stringyear & " and [CityID]='" & PopUpValue1 & "'", CON
    If RS.EOF = False Then
       txtDist = RS(0)
       txtState = RS(1)
    End If
    
    txtPin.SetFocus
End If

PopUpValue1 = ""
PopUpValue2 = ""

End Sub

Private Sub Form_Activate()
cmdAdd_1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub



Private Sub txtcity_GotFocus()
HIT
If PopUpValue1 <> "" Then
   
    txtCity = PopUpValue2
    txtCityID = PopUpValue1
    
    If RS.State = 1 Then RS.close
    RS.Open "select [District],[State] FROM  [CityView] " & _
    "where " & stringyear & " and [CityID]='" & PopUpValue1 & "'", CON
    If RS.EOF = False Then
       txtDist = RS(0)
       txtState = RS(1)
    End If
    
    txtPin.SetFocus
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
Private Sub Form_Load()
 Me.Top = 1000
 Me.Left = 500
 
 addDis1
 
 BackColorFrom Me
 
End Sub
Sub addDis2()

ListDis2.Clear

If RS.State = 1 Then RS.close
RS.Open "Select districtname from  districts  where agentname='" & txtname & "' and " & stringyear & "", CON, adOpenStatic, adLockPessimistic
If RS.RecordCount > 0 Then
    ListDis2.Clear
    RS.MoveFirst
    While Not RS.EOF
         If IsNull(RS(0)) = False Then
               ListDis2.AddItem RS(0)
         End If
         If Not RS.EOF Then
                RS.MoveNext
         End If
   Wend
   
   ListDis2.Selected(0) = True
End If


End Sub
Sub addDis1()

ListDis1.Clear

If RS.State = 1 Then RS.close
RS.Open "Select districtname from  districts  where (len(agentname)<=1 or agentname is null) and " & stringyear & "", CON, adOpenStatic, adLockPessimistic
If RS.RecordCount > 0 Then
    ListDis1.Clear
    RS.MoveFirst
    While Not RS.EOF
         If IsNull(RS(0)) = False Then
               ListDis1.AddItem RS(0)
         End If
         If Not RS.EOF Then
                RS.MoveNext
         End If
   Wend
   
   ListDis1.Selected(0) = True
End If


End Sub

Private Sub Removecd_Click()
If ListDis2.ListCount > 0 Then
    Dim I
    Dim delitem As Integer
    For I = 0 To ListDis2.ListCount - 1
        If ListDis2.Selected(I) Then
                ListDis1.AddItem ListDis2.List(I)
                delitem = I
         End If
    Next
    ListDis2.RemoveItem delitem
End If

End Sub
