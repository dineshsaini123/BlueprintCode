VERSION 5.00
Begin VB.Form frmCollegeList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "College LIst"
   ClientHeight    =   2892
   ClientLeft      =   3960
   ClientTop       =   2520
   ClientWidth     =   4860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2892
   ScaleWidth      =   4860
   Begin VB.ComboBox cboCat 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   180
      Width           =   3300
   End
   Begin VB.CommandButton cmdSticker 
      Caption         =   "&Sticker Print"
      Height          =   450
      Left            =   1590
      TabIndex        =   4
      Top             =   1755
      Width           =   2010
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   450
      Left            =   1590
      TabIndex        =   3
      Top             =   2295
      Width           =   2010
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print College List"
      Height          =   450
      Left            =   1590
      TabIndex        =   2
      Top             =   1230
      Width           =   2010
   End
   Begin VB.ComboBox cboCity 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   660
      Width           =   3300
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "School Category :"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   210
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "City :"
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   690
      Width           =   765
   End
End
Attribute VB_Name = "frmCollegeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset

Private Sub cmdexit_Click()
  Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    s = s & " and {ISSUEBOOK.fyear}='" & session & "' and {ISSUEBOOK.setupid}=" & setupid & ""
    
    DSNNew
    
    frmAgentLadger.cr.Reset
    frmAgentLadger.cr.ReportFileName = rptPath & "/CollegeList.rpt"
    frmAgentLadger.cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If cboCity.Text <> "" And cboCat.Text <> "" Then
       frmAgentLadger.cr.ReplaceSelectionFormula "{college.fyear}='" & session & "' and {college.setupid}=" & setupid & " and {college.district}='" & cboCity.Text & "' and {college.states}='" & cboCat.Text & "'"
    ElseIf cboCity.Text <> "" And cboCat.Text = "" Then
       frmAgentLadger.cr.ReplaceSelectionFormula "{college.fyear}='" & session & "' and {college.setupid}=" & setupid & " and {college.district}='" & cboCity.Text & "'"
    Else
       frmAgentLadger.cr.ReplaceSelectionFormula "{college.fyear}='" & session & "' and {college.setupid}=" & setupid & " and {college.states}='" & cboCat.Text & "'"
    End If
    
    frmAgentLadger.cr.WindowShowPrintBtn = True
    frmAgentLadger.cr.WindowState = crptMaximized
    frmAgentLadger.cr.WindowShowPrintBtn = True
    frmAgentLadger.cr.WindowShowPrintSetupBtn = True
    frmAgentLadger.cr.WindowShowSearchBtn = True

    frmAgentLadger.cr.Action = 1
 
End Sub
Private Sub cmdSticker_Click()

    DSNNew

    frmAgentLadger.cr.Reset
    frmAgentLadger.cr.ReportFileName = rptPath & "/Sticker.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    
    If cboCity.Text <> "" And cboCat.Text <> "" Then
       frmAgentLadger.cr.ReplaceSelectionFormula "{college.fyear}='" & session & "' and {college.setupid}=" & setupid & " and {college.district}='" & cboCity.Text & "' and {college.states}='" & cboCat.Text & "'"
    ElseIf cboCity.Text <> "" And cboCat.Text = "" Then
       frmAgentLadger.cr.ReplaceSelectionFormula "{college.fyear}='" & session & "' and {college.setupid}=" & setupid & " and {college.district}='" & cboCity.Text & "'"
    Else
       frmAgentLadger.cr.ReplaceSelectionFormula "{college.fyear}='" & session & "' and {college.setupid}=" & setupid & " and {college.states}='" & cboCat.Text & "'"
    End If
    
    frmAgentLadger.cr.WindowShowPrintBtn = True
    frmAgentLadger.cr.WindowState = crptMaximized
    frmAgentLadger.cr.WindowShowPrintBtn = True
    frmAgentLadger.cr.WindowShowPrintSetupBtn = True
    frmAgentLadger.cr.WindowShowSearchBtn = True

    frmAgentLadger.cr.Action = 1

End Sub

Private Sub Form_Load()

If RS.State = 1 Then RS.close
RS.Open "select distinct(district) from college where " & stringyear, con
While RS.EOF = False
   cboCity.AddItem RS.Fields(0).value
   RS.MoveNext
Wend

If RS.State = 1 Then RS.close
RS.Open "select distinct(BookCategory) from BookCategory", con
While RS.EOF = False
   cboCat.AddItem RS.Fields(0).value
   RS.MoveNext
Wend


BackColorFrom Me

End Sub
