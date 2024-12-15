VERSION 5.00
Begin VB.Form frmProdust 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Master"
   ClientHeight    =   7170
   ClientLeft      =   7440
   ClientTop       =   1875
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   4920
   Begin VB.OptionButton NonConsume 
      Caption         =   "Non Consume"
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   4920
      TabIndex        =   9
      Top             =   810
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.OptionButton consume 
      Caption         =   "Consume"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   390
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   30
      TabIndex        =   5
      Top             =   5970
      Width           =   4785
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3585
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1275
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2430
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.TextBox txtproduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4875
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FF0000&
      Height          =   5550
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   4845
   End
   Begin VB.Label auto 
      Height          =   420
      Left            =   5040
      TabIndex        =   10
      Top             =   1620
      Width           =   375
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
      Left            =   3510
      TabIndex        =   6
      Top             =   6870
      Width           =   1335
   End
End
Attribute VB_Name = "frmProdust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bb As Boolean
Private Sub cmdDel_Click()
'On Error Resume Next

On Error GoTo aa:

If rs.State = 1 Then rs.Close
         rs.Open "select * from ProductMaster where Name='" & List1 & "'", CON, adOpenDynamic, adLockOptimistic
         If rs.EOF = False Then
         If MsgBox("Do U Want To Delete ?", vbInformation + vbYesNo, "Message") = vbYes Then
            rs.Delete
            addProduct
            txtproduct.Text = ""
         End If
         End If
         
 Exit Sub
aa:
 MsgBox "" & err.DESCRIPTION
         
         
         
End Sub

Private Sub cmdMain_Click()
Unload Me
End Sub
Sub addProduct()
'On Error Resume Next
      Set rs = New ADODB.Recordset
         rs.Open "select * from  ProductMaster order by name", CON, adOpenDynamic, adLockOptimistic
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
       
On Error GoTo savedata:
       
   If txtproduct.Text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If
            
      If rs.State = 1 Then rs.Close
         
         If auto.Caption = "" Then
         If rs.State = 1 Then rs.Close
         rs.Open "select * from ProductMaster where name='" & txtproduct.Text & "'", CON, adOpenDynamic, adLockOptimistic
         Else
         rs.Open "select * from ProductMaster where auto=" & auto.Caption & "", CON, adOpenDynamic, adLockOptimistic
         End If
         If rs.EOF = True Then
            rs.addNew
            rs.Fields(0).Value = txtproduct.Text
            If consume.Value = True Then
               rs.Fields(1).Value = consume.Caption
               rs.Fields(0).Value = txtproduct.Text
            Else
               rs.Fields(0).Value = txtproduct.Text
               rs.Fields(1).Value = NonConsume.Caption
            End If
            
            rs.Update
            txtproduct.Text = ""
            addProduct
          Else
            rs.Fields(0).Value = txtproduct.Text
            If consume.Value = True Then
               rs.Fields(1).Value = consume.Caption
               Else
               rs.Fields(1).Value = NonConsume.Caption
            End If
            
            
            rs.Update
            txtproduct.Text = ""
            addProduct
         End If


Exit Sub

savedata:

MsgBox "" & err.DESCRIPTION

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

If rs.State = 1 Then rs.Close
rs.Open "select * from ProductMaster where Name='" & txtproduct & "'", CON, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
auto.Caption = rs!auto
If rs.Fields(1).Value = "Consume" Then
   consume.Value = True
Else
   NonConsume.Value = True
End If
End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
   On Error Resume Next
     
   If txtproduct.Text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If

   If KeyAscii = 13 Then
      If rs.State = 1 Then rs.Close
         rs.Open "select * from ProductMaster where Name='" & txtproduct & "'", CON, adOpenDynamic, adLockOptimistic
         If rs.EOF = True Then
            rs.addNew
            rs.Fields(1).Value = txtproduct.Text
            If consume.Value = True Then
               rs.Fields(1).Value = consume.Caption
               Else
               rs.Fields(2).Value = NonConsume.Caption
            End If
            
            
            rs.Update
            txtproduct.Text = ""
            addProduct
          Else
            rs.Fields(1).Value = txtproduct.Text
            
            If consume.Value = True Then
               rs.Fields(1).Value = consume.Caption
               Else
               rs.Fields(2).Value = NonConsume.Caption
            End If
            
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


