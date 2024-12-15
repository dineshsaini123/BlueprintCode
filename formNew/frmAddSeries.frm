VERSION 5.00
Begin VB.Form frmAddSeries 
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13725
   Icon            =   "frmAddSeries.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   13725
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox txtseries 
      Height          =   315
      Left            =   7245
      TabIndex        =   6
      Top             =   360
      Width           =   6405
   End
   Begin VB.CommandButton Command1_ok 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   555
      Left            =   6450
      TabIndex        =   5
      Top             =   1365
      Width           =   525
   End
   Begin VB.TextBox txtsearch 
      Height          =   375
      Left            =   75
      TabIndex        =   3
      Top             =   360
      Width           =   6315
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   555
      Left            =   6450
      TabIndex        =   2
      Top             =   765
      Width           =   525
   End
   Begin VB.ListBox List2_search 
      Height          =   6810
      Left            =   75
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   780
      Width           =   6315
   End
   Begin VB.ListBox List1_add 
      Height          =   6885
      Left            =   7245
      TabIndex        =   0
      Top             =   720
      Width           =   6360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Series Name"
      Height          =   195
      Left            =   7290
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmAddSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
   For J = 0 To List2_search.ListCount - 1
      If List2_search.Selected(J) = True Then
         List1_add.AddItem List2_search.List(J)
      End If
   Next
   Command1_ok.Enabled = True
End Sub

Private Sub Command1_ok_Click()
    For J = 0 To List1_add.ListCount - 1
       con.Execute "update books set sername='" & UCase(txtseries) & "' where bookname='" & List1_add.List(J) & "'"
    Next
    
    MsgBox "Data Saved...", vbInformation
    
End Sub

Private Sub Form_Load()

Me.Top = 800
Me.Left = 300

If RS.State = 1 Then RS.close
''RS.Open "select distinct SerName from BOOKS order by SerName", con
RS.Open "select distinct Series from SeriesMaster order by Series", con
While RS.EOF = False
If Not IsNull(RS(0)) Then
txtseries.AddItem RS(0)
End If
RS.MoveNext
Wend

End Sub

Private Sub txtsearch_Change()
  
   List2_search.Clear

   If RS.State = 1 Then RS.close
   RS.Open "select BOOKNAME from BOOKS where BOOKNAME like '" & txtsearch & "%'", con
   While RS.EOF = False
     List2_search.AddItem RS(0)
     RS.MoveNext
   Wend
   
End Sub
