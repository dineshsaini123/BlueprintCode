VERSION 5.00
Begin VB.Form frmIssuerpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agent Wise Issue"
   ClientHeight    =   3504
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4824
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3504
   ScaleWidth      =   4824
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   696
      ItemData        =   "frmIsuueRpt.frx":0000
      Left            =   585
      List            =   "frmIsuueRpt.frx":000D
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   945
      Width           =   1635
   End
   Begin VB.ComboBox cboCity 
      Height          =   315
      Left            =   600
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   3900
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2415
      Width           =   2010
   End
   Begin VB.CommandButton cmdSticker 
      Caption         =   "&Print"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1875
      Width           =   2010
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Name"
      Height          =   240
      Left            =   600
      TabIndex        =   3
      Top             =   300
      Width           =   1725
   End
End
Attribute VB_Name = "frmIssuerpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset

Private Sub cmdexit_Click()
  Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    DSNNew
    
    frmAgentLadger.cr.Reset
    frmAgentLadger.cr.ReportFileName = rptPath & "/CollegeList.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If cboCity.Text <> "" Then
     frmAgentLadger.cr.ReplaceSelectionFormula "{college.district}='" & cboCity.Text & "'"
    End If
    frmAgentLadger.cr.WindowShowPrintBtn = True
    frmAgentLadger.cr.WindowState = crptMaximized
    frmAgentLadger.cr.WindowShowPrintBtn = True
    frmAgentLadger.cr.WindowShowPrintSetupBtn = True
    frmAgentLadger.cr.WindowShowSearchBtn = True

    frmAgentLadger.cr.Action = 1
 
End Sub

Private Sub cmdSticker_Click()
    
    Dim s As String
    
    s = "{info.fyear}='" & session & "' and {info.setupid}=" & setupid & ""
    
    For I = 0 To List1.ListCount - 1
    If List1.Selected(I) = True Then
    If s = "" Then
    s = "left(ucase(trim({info.vno})),1)='" & UCase(List1.List(I)) & "'"
    Else
    s = s & " or " & "left(trim(ucase({info.vno})),1)='" & UCase(List1.List(I)) & "'"
    End If
    End If
    Next
    
    DSNNew
    
    frmAgentLadger.cr.Reset
    frmAgentLadger.cr.ReportFileName = rptPath & "/AgentWiseIssueBook.rpt"
    frmAgentLadger.cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    frmAgentLadger.cr.ReplaceSelectionFormula s
    If cboCity.Text <> "" Then
    If s <> "" Then
     frmAgentLadger.cr.ReplaceSelectionFormula s & " And " & "{info.aname}='" & cboCity.Text & "'"
    Else
     frmAgentLadger.cr.ReplaceSelectionFormula "{info.aname}='" & cboCity.Text & "'"
    End If
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
   RS.Open "select distinct(aname) from Info where " & stringyear, con
   While RS.EOF = False
      cboCity.AddItem RS.Fields(0).value
      RS.MoveNext
   Wend
   
   BackColorFrom Me
   
End Sub

