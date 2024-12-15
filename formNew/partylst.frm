VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form partylist 
   Caption         =   "Party Address Details "
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List_option 
      Height          =   2535
      Left            =   1740
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   1260
      Width           =   3990
   End
   Begin VB.ComboBox cbostate 
      Height          =   315
      Left            =   1740
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   4005
   End
   Begin VB.ComboBox Combosldistrictcode 
      Height          =   315
      Left            =   1740
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   4005
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   585
      Left            =   1755
      TabIndex        =   1
      Top             =   4020
      Width           =   1545
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   585
      Left            =   3375
      TabIndex        =   0
      Top             =   4020
      Width           =   1455
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   180
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Party Label"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   255
      Left            =   615
      TabIndex        =   5
      Top             =   420
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "District Name"
      Height          =   195
      Left            =   630
      TabIndex        =   3
      Top             =   900
      Width           =   1035
   End
End
Attribute VB_Name = "partylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbostate_Click()


If cboState.Text = "" Then Exit Sub
Me.Combosldistrictcode.Clear
If RS.State = 1 Then RS.close
RS.Open "select distinct distcode from sledger where " & stringyear & " and states='" & cboState.Text & "'", con, adOpenDynamic, adLockReadOnly, adCmdText
If Not RS.EOF Then
Do While Not RS.EOF

If Not IsNull(RS(0)) Then
    Me.Combosldistrictcode.AddItem RS(0)
End If
If Not RS.EOF Then
    RS.MoveNext
End If

Loop
End If
RS.close


End Sub

Private Sub Commandreturn_Click()
Unload Me
End Sub
Private Sub Commandshow_Click()

Dim ss1

ss1 = ""

For I = 0 To List_option.ListCount - 1
   If List_option.Selected(I) = True Then
      If ss1 = "" Then
        ss1 = List_option.List(I) & " = " & -1
      Else
        ss1 = ss1 & " or " & List_option.List(I) & " = " & -1
      End If
   End If
Next


con.Execute "Delete  from tmps_ledger where " & stringyear & " and len(SUBLEDGER)>0"
's1 = App.Path & "\" & "2003-04" & "\tchitra.mdb"

If Combosldistrictcode.Text <> "" Then

If ss1 <> "" Then
   con.Execute "insert into  tmps_ledger  select * from sledger where " & stringyear & " and distcode= '" & Combosldistrictcode.Text & "' and gledger='SUNDRY DEBTORS' and " & ss1
Else
   con.Execute "insert into  tmps_ledger  select * from sledger where " & stringyear & " and distcode= '" & Combosldistrictcode.Text & "' and gledger='SUNDRY DEBTORS'"
End If

Else

If ss1 <> "" Then
 con.Execute "insert into  tmps_ledger  select * from sledger where " & stringyear & " and gledger='SUNDRY DEBTORS' and states='" & cboState.Text & "' and " & ss1
Else
 con.Execute "insert into  tmps_ledger  select * from sledger where " & stringyear & " and gledger='SUNDRY DEBTORS' and states='" & cboState.Text & "'"
End If

End If

DoEvents
DoEvents
DoEvents
DoEvents
DoEvents


DSNNew

If MsgBox("Print ?", vbInformation + vbYesNo) = vbYes Then

 cr1.Reset
 'cr1.ReportFileName = App.Path & "\2003-04\PartyLabel.rpt"
 cr1.ReportFileName = rptPath & "\PartyLabel.rpt"
 cr1.WindowState = crptMaximized
 cr1.WindowShowRefreshBtn = True
 cr1.Action = 1

End If




End Sub

Private Sub Form_Load()

Dim RS As New ADODB.Recordset
 
 RS.Open "select * from DISTRICTS where " & stringyear & "  order by DISTRICTNAME", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
            Me.Combosldistrictcode.AddItem RS!DISTRICTNAME
          
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    
    
    

If RS.State = 1 Then RS.close
RS.Open "select distinct states from sledger where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
If Not RS.EOF Then
    Do While Not RS.EOF
      If Not IsNull(RS(0)) Then
         Me.cboState.AddItem RS(0)
      End If
      
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If
RS.close
    
    
    
'==============
List_option.Clear

If RS.State = 1 Then RS.close
RS.Open "select * from sledger where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
For I = 25 To RS.Fields.Count - 1
   List_option.AddItem RS.Fields(I).Name
Next
    
    
    
    
End Sub
