VERSION 5.00
Begin VB.Form frmReplaceName 
   Caption         =   "Replace Name"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5256
   Icon            =   "frmReplaceName.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5256
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbocity 
      Height          =   288
      Left            =   1068
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   3492
   End
   Begin VB.CheckBox Check1_citywise 
      Caption         =   "City Wise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1116
      TabIndex        =   5
      Top             =   216
      Width           =   1632
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Replace Rep. Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1116
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2856
      Width           =   3048
   End
   Begin VB.ComboBox cborep1 
      Height          =   288
      Left            =   1080
      TabIndex        =   1
      Top             =   1212
      Width           =   3492
   End
   Begin VB.ComboBox cborep2 
      Height          =   288
      Left            =   1080
      TabIndex        =   0
      Top             =   2232
      Width           =   3492
   End
   Begin VB.Label lblCity 
      BackStyle       =   0  'Transparent
      Caption         =   "City :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   108
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   1860
      Picture         =   "frmReplaceName.frx":000C
      Top             =   1704
      Width           =   384
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Rep. Name  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1212
      Width           =   1092
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Rep. Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   2232
      Width           =   1152
   End
End
Attribute VB_Name = "frmReplaceName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_citywise_Click()

If Check1_citywise.value = 1 Then
   lblCity(0).Visible = True
   cbocity.Visible = True
Else
   lblCity(0).Visible = False
   cbocity.Visible = False

End If


End Sub

Private Sub cmdedit_Click()
 
 
If Check1_citywise.value = 0 Then
 
 con.Execute "update SLEDGER set RepName1='" & cborep2.text & "' where repname1='" & cborep1.text & "'"
 con.Execute "update SLEDGER set RepName2='" & cborep2.text & "' where repname2='" & cborep1.text & "'"
 con.Execute "update SLEDGER set RepName3='" & cborep2.text & "' where repname3='" & cborep1.text & "'"
 con.Execute "update SLEDGER set RepName4='" & cborep2.text & "' where repname4='" & cborep1.text & "'"
 con.Execute "update SLEDGER set RepName5='" & cborep2.text & "' where repname5='" & cborep1.text & "'"
 con.Execute "update SLEDGER set RepName6='" & cborep2.text & "' where repname6='" & cborep1.text & "'"
 
Else

If (cbocity.text <> "") Then

 con.Execute "update SLEDGER set RepName1='" & cborep2.text & "' where (repname1='" & cborep1.text & "' and ADDRESS3='" & cbocity.text & "')"
 con.Execute "update SLEDGER set RepName2='" & cborep2.text & "' where (repname2='" & cborep1.text & "' and ADDRESS3='" & cbocity.text & "')"
 con.Execute "update SLEDGER set RepName3='" & cborep2.text & "' where (repname3='" & cborep1.text & "' and ADDRESS3='" & cbocity.text & "')"
 con.Execute "update SLEDGER set RepName4='" & cborep2.text & "' where (repname4='" & cborep1.text & "' and ADDRESS3='" & cbocity.text & "')"
 con.Execute "update SLEDGER set RepName5='" & cborep2.text & "' where (repname5='" & cborep1.text & "' and ADDRESS3='" & cbocity.text & "')"
 con.Execute "update SLEDGER set RepName6='" & cborep2.text & "' where (repname6='" & cborep1.text & "' and ADDRESS3='" & cbocity.text & "')"

Else
 
 MsgBox "Select City...", vbInformation
 Exit Sub

End If

 
End If
 
 MsgBox "Changed Successfuuly...", vbInformation

End Sub
Private Sub Form_Load()

If RS.State = 1 Then RS.close
RS.Open "select Rep as Representative from SalesRepQry order by Rep", CON_blue
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cborep1.AddItem RS(0)
        Me.cborep2.AddItem RS(0)
       End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If


If RS.State = 1 Then RS.close
RS.Open "select distinct address3 from Sledger where len(address3)>0 and gledger='SUNDRY DEBTORS' order by address3", con
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cbocity.AddItem RS(0)
       End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If



End Sub
