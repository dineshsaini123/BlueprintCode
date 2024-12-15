VERSION 5.00
Begin VB.Form StateTaxList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "State Tax List"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   435
      Left            =   -780
      TabIndex        =   8
      Top             =   2940
      Width           =   675
   End
   Begin VB.TextBox txtLess 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   1860
      Width           =   1335
   End
   Begin VB.TextBox txtAdd 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   1620
      TabIndex        =   3
      Top             =   2940
      Width           =   975
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   660
      TabIndex        =   2
      Top             =   2940
      Width           =   855
   End
   Begin VB.ListBox StateList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   300
      Width           =   2895
   End
   Begin VB.ComboBox with_without 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "StateTaxList.frx":0000
      Left            =   300
      List            =   "StateTaxList.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Select a method"
      Top             =   420
      Width           =   2535
   End
   Begin VB.Label Less_Label 
      AutoSize        =   -1  'True
      Caption         =   "Less :"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   420
   End
   Begin VB.Label Add_Label 
      AutoSize        =   -1  'True
      Caption         =   "Add :"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1380
      Width           =   375
   End
End
Attribute VB_Name = "StateTaxList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
''addnewst = InputBox("Enter State Name", "Adding to list")
''If Not Len(addnewst) = 0 Then
''addwithadd = InputBox("Enter value of add field/with C form", "Add-With")
''addwithless = InputBox("Enter value of less field/with C form", "Less-With")
''addwithoutadd = InputBox("Enter value of add field/without C form", "Less-With")
''addwithoutless = InputBox("Enter value of less field/without C form", "Less-With")
''
''
''
''If rs.State = 1 Then rs.Close
''    rs.Open "select * from state_tax_list", CON, adOpenDynamic, adLockOptimistic
''    rs.AddNew
''    rs!statename = addnewst
''    rs!with_without = "with"
''    rs!add_val = addwithadd
''    rs!less_val = addwithless
''    rs.Update
''    rs.AddNew
''     rs!statename = addnewst
''    rs!with_without = "without"
''    rs!add_val = addwithoutadd
''    rs!less_val = addwithoutless
''    rs.Update
''    rs.Close
''
''MsgBox "Data Added", vbInformation
''StateList.AddItem (addnewst)
''
''End If
End Sub

Private Sub btnOK_Click()
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from state_tax_list where statename='" & StateList.List(StateList.ListIndex) & "' and with_without = '" & with_without.Text & "'", CON, adOpenDynamic, adLockOptimistic
    If rs.EOF = True Then
        rs.AddNew
        rs![add_val] = txtAdd
        rs![less_val] = txtLess
        rs!with_without = with_without.Text
        rs!statename = StateList.List(StateList.ListIndex)
        rs.Update
    Else
        rs![add_val] = txtAdd
        rs![less_val] = txtLess
        rs!with_without = with_without.Text
        rs!statename = StateList.List(StateList.ListIndex)
        rs.Update
    End If
        
    MsgBox "Date Saved ...", vbInformation
        
        
''    ElseIf with_without.Text = "Without C Form" Then
''    If rs.State = 1 Then rs.Close
''        rs.Open "select add_val,less_val from state_tax_list where statename='" & StateList.List(StateList.ListIndex) & "' and with_without = 'without'", CON, adOpenDynamic, adLockOptimistic
''        If Not rs.EOF Then
''
''        rs(0) = txtAdd
''        rs(1) = txtLess
''        rs.Update
''        rs.Close
''
''        End If
''
''    End If
  
End Sub

Private Sub Form_Load()

If rs.State = 1 Then rs.Close
rs.Open "select [state] from [state] order by [state]", CON, adOpenDynamic, adLockReadOnly
If Not rs.EOF Then
Dim i As Integer
i = 0
While Not rs.EOF
StateList.AddItem rs(0)
rs.MoveNext
Wend

End If


with_without.ListIndex = 0


End Sub


Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub StateList_Click()
If rs.State = 1 Then rs.Close
rs.Open "select add_val,less_val from state_tax_list where statename='" & StateList.List(StateList.ListIndex) & "' and with_without = '" & with_without.Text & "'", CON
If rs.EOF = False Then
  txtAdd = rs(0)
  txtLess = rs(1)
Else
  txtAdd = 0
  txtLess = 0

End If

End Sub

Private Sub with_without_Click()

If rs.State = 1 Then rs.Close
rs.Open "select add_val,less_val from state_tax_list where statename='" & StateList.List(StateList.ListIndex) & "' and with_without = '" & with_without.Text & "'", CON
If rs.EOF = False Then
  txtAdd = rs(0)
  txtLess = rs(1)
Else
  txtAdd = 0
  txtLess = 0

End If

End Sub


