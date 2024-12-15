VERSION 5.00
Begin VB.Form UnitMaster 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Unit Creation"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FF0000&
      Height          =   5100
      Left            =   30
      TabIndex        =   6
      Top             =   360
      Width           =   3975
   End
   Begin VB.TextBox txtproduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   5490
      Width           =   3915
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   495
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   495
         Left            =   1035
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   495
         Left            =   2865
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E98A0A&
      BackStyle       =   0  'Transparent
      Caption         =   "Esc To Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2610
      TabIndex        =   7
      Top             =   6390
      Width           =   1335
   End
End
Attribute VB_Name = "UnitMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bb As Boolean
Private Sub cmdDel_Click()
On Error Resume Next
If rs.State = 1 Then rs.Close
         rs.Open "select * from UnitMaster where Name='" & List1 & "'", CON, adOpenDynamic, adLockOptimistic
         If rs.EOF = False Then
         If MsgBox("Do U Want To Delete ?", vbInformation + vbYesNo, "Message") = vbYes Then
            rs.Delete
            addProduct
            txtproduct.Text = ""
         End If
         End If
         
End Sub

Private Sub cmdMain_Click()
Unload Me
End Sub
Sub addProduct()
If rs.State = 1 Then rs.Close
rs.Open "select * from  UnitMaster", CON, adOpenDynamic, adLockOptimistic
List1.Clear

If rs.EOF = False Then
   While rs.EOF = False
       List1.AddItem rs.Fields(0).Value
       rs.MoveNext
   Wend
End If
End Sub

Private Sub cmdRef_Click()
 txtproduct.Text = ""
End Sub

Private Sub cmdSave_Click()
 save
End Sub
Sub save()
       
On Error GoTo ss:
       
   If txtproduct.Text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If

            
      If rs.State = 1 Then rs.Close
          rs.Open "select * from UnitMaster where Name='" & txtproduct & "'", CON, adOpenDynamic, adLockOptimistic
         If rs.EOF = True Then
            rs.addNew
            rs.Fields(0).Value = txtproduct.Text
            rs.Update
            txtproduct.Text = ""
            addProduct
          Else
            rs.Fields(0).Value = txtproduct.Text
            rs.Update
            txtproduct.Text = ""
            addProduct
         End If

Exit Sub
ss:
MsgBox err.DESCRIPTION



End Sub

Private Sub Form_Activate()
txtproduct.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    SendKeys "{tab}"
 End If
 If KeyCode = 27 Then
    Unload Me
 End If
End Sub
Private Sub Form_Load()
addProduct
frmProdust.Left = 2500
'Call frmBackColor(frmProdust)


'Call UserPermission(cmdSave, cmdDel, cmdRef)

End Sub
Private Sub List1_Click()
txtproduct.Text = List1.Text
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
   On Error Resume Next
     
   If txtproduct.Text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If

   If KeyAscii = 13 Then
      If rs.State = 1 Then rs.Close
         rs.Open "select * from UnitMaster where Name='" & txtproduct & "'", CON, adOpenDynamic, adLockOptimistic
         If rs.EOF = True Then
            rs.addNew
            rs.Fields(1).Value = txtproduct.Text
            rs.Update
            txtproduct.Text = ""
            addProduct
          Else
            rs.Fields(1).Value = txtproduct.Text
            rs.Update
            txtproduct.Text = ""
            addProduct
         End If
   End If

End Sub

Private Sub List2_Click()
txtproduct.Text = List1.Text
End Sub
Private Sub txtproduct_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      save
   End If
End Sub

