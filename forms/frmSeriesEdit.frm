VERSION 5.00
Begin VB.Form frmSeriesEdit 
   ClientHeight    =   1812
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6780
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1812
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdchange 
      Caption         =   "&Change"
      Height          =   555
      Left            =   3540
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtserTo 
      Height          =   375
      Left            =   3540
      TabIndex        =   2
      Top             =   420
      Width           =   3075
   End
   Begin VB.ComboBox cboser 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "New Series Name"
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "To "
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Series Name"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
End
Attribute VB_Name = "frmSeriesEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdchange_Click()
If MsgBox("Want to Change ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "update BOOKS set sername='" & Trim(txtserTo) & "' where sername='" & cboser.Text & "'"
End If
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.close
'RS.Open "select distinct SerName from BOOKS order by SerName", con
RS.Open "select distinct series from SeriesMaster order by series", con
While RS.EOF = False
cboser.AddItem RS(0) & ""
RS.MoveNext
Wend

End Sub
