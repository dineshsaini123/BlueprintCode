VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form bankadvice 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   1410
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtadno 
      Height          =   345
      Left            =   3690
      TabIndex        =   3
      Top             =   135
      Width           =   1845
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   1440
      Left            =   45
      TabIndex        =   8
      Top             =   2205
      Visible         =   0   'False
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   2540
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   20000
      TextRTF         =   $"badvice.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   960
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FromPage        =   1
      PrinterDefault  =   0   'False
      ToPage          =   1
   End
   Begin VB.CommandButton print1 
      Height          =   345
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3900
      TabIndex        =   5
      Text            =   "100 %"
      Top             =   4710
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   630
      Left            =   2295
      Picture         =   "badvice.frx":0080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1125
      Width           =   1605
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   630
      Left            =   3960
      Picture         =   "badvice.frx":0C64
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1125
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   3690
      TabIndex        =   0
      Top             =   630
      Width           =   1845
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Advice No.(If do you want to start own advice no the use this box)"
      Height          =   435
      Left            =   540
      TabIndex        =   9
      Top             =   135
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No."
      Height          =   285
      Left            =   2340
      TabIndex        =   7
      Top             =   630
      Width           =   1305
   End
End
Attribute VB_Name = "bankadvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLSTRING As String
Dim RS As ADODB.Recordset

Private Sub Combo1_Change()
If Trim(Combo1.Text) = "75 %" Then
    r1.Font.Size = 8
End If
If Trim(Combo1.Text) = "100 %" Then
    r1.Font.Size = 10
End If
If Trim(Combo1.Text) = "200 %" Then
    r1.Font.Size = 18
End If
If Trim(Combo1.Text) = "125 %" Then
    r1.Font.Size = 12
End If
If Trim(Combo1.Text) = "150 %" Then
    r1.Font.Size = 14
End If
End Sub

Private Sub Combo1_Click()
'r1.row = 1
If Trim(Combo1.Text) = "75 %" Then
    r1.Font.Size = 8
End If
If Trim(Combo1.Text) = "100 %" Then
    r1.Font.Size = 10
End If
If Trim(Combo1.Text) = "200 %" Then
    r1.Font.Size = 18
End If
If Trim(Combo1.Text) = "125 %" Then
    r1.Font.Size = 12
End If
If Trim(Combo1.Text) = "150 %" Then
    r1.Font.Size = 14
End If


End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Commandreturn_Click()
Unload Me
End Sub

Private Sub Commandshow_Click()
If Val(Text1) > 0 Then
   genreport
  
Else
   MsgBox "Invalid Invoice No."
   Text1.SetFocus
End If
    
End Sub

Public Function printnow()
    Dim X As Long
    Dim p As Printer
    For I = 0 To Printers.Count - 1
        Set p = Printers(I)
        If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.Text) Then
            Exit For
        End If
    Next
    For I = 1 To (Printdlg.UpDown1.value)
        X = Shell("" + VB.App.Path + "\ppp.bat " + VB.App.Path + "\vipin.txt " & Trim(p.Port))
    Next
    Printdlg.UpDown1.value = 1
    Printdlg.Text1.Text = "1"
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}"
End If

End Sub

Private Sub Form_Load()
  Dim X
  X = Left("dfdfg", 2)
  Combo1.Top = r1.Top + r1.Height + 30
  Me.Top = 0
  Me.Left = 0
  Set RS = New ADODB.Recordset
  Dim rs1 As New ADODB.Recordset
  rs1.Open "select * from setup1 where " & stringyear, CON, adOpenDynamic, adLockOptimistic
  If rs1.RecordCount > 0 Then
     txtadno.Text = rs1!bankadviceno
  End If
  
 BackColorFrom Me

End Sub
Private Sub Form_Resize()
    If Me.Width > 350 And Me.Height > 1500 Then
        r1.Width = Me.Width - 250
        r1.Height = Me.Height - 1000
    '    Command1.Top = r1.Top + r1.Height + 30
        Combo1.Top = r1.Top + r1.Height + 30
        export.Top = Combo1.Top
        Me.print1.Top = export.Top
    End If
End Sub

Private Sub return1_Click()
    Unload Me
End Sub

Private Sub print_Click()

End Sub
Function genreport()



  Dim rs1 As ADODB.Recordset
   Set rs1 = New ADODB.Recordset
   Set RS = New ADODB.Recordset
   
   
   If RS.State = 1 Then RS.close
      RS.Open "select * from invoicea where " & stringyear & " and advno=" + txtadno.Text + "", CON, adOpenDynamic, adLockOptimistic, adCmdText
      If RS.RecordCount > 0 Then
         If MsgBox("Advice allready exist....", vbOKCancel) = vbCancel Then
             Exit Function
         End If
      End If
    
    If RS.State = 1 Then RS.close
   RS.Open "select * from invoicea where " & stringyear & " and invoiceno=" + Text1.Text + "", CON, adOpenDynamic, adLockOptimistic, adCmdText
   If RS.RecordCount > 0 Then
        If RS!baa > 0 Then
            If RS!advno > 0 Then
               If MsgBox("Advice all ready printed....", vbOKCancel) = vbCancel Then
                  Exit Function
               Else
                  'txtadno.Text = rs!advno
               End If
            End If
            
            RS!advno = Val(txtadno.Text)
            RS.update
            If RS.BOF Then
               MsgBox "Invoice not found"
               Exit Function
            End If
            rs1.Open "select * from setup1 where " & stringyear, CON, adOpenDynamic, adLockOptimistic, adCmdTable
            rs1.MoveFirst
           
            Open "" + VB.App.Path + "\vipin.txt" For Output As #1
            Print #1, ""; Chr(27) + Chr(15)
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, "                " & txtadno.Text; Tab(48 + 35); RS!invoicedate
            rs1!bankadviceno = txtadno.Text + 1 'rs1!bankadviceno + 1
            txtadno.Text = txtadno + 1
            rs1.update
            Print #1, ""
            Print #1, Tab(18 + 15); Trim(RS!through)
            Print #1, Tab(18 + 15); Trim(RS!through1)
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, Tab(20 + 15); Trim(RS!biltyno); Tab(48 + 35); IIf(IsNull(RS!BILTYDATE), " / /   ", RS!BILTYDATE)
            Print #1, ""
            Print #1, Tab(20 + 15); Trim(RS!INVOICENO); Tab(48 + 35); IIf(IsNull(RS!invoicedate), " / /   ", RS!invoicedate)
            If rs1.State = 1 Then rs1.close
            rs1.Open "select * from sledger where " & stringyear & " and gledger='SUNDRY DEBTORS' and subledger='" + Trim(RS!SUBLEDGER) + "'", CON, adOpenDynamic, adLockReadOnly
            Print #1, ""
            Print #1, ""
            If Not rs1.BOF Then
                Print #1, Tab(20 + 15); IIf(IsNull(rs1!DESCFORINVOICE), "", rs1!DESCFORINVOICE)
                Print #1, Tab(20 + 15); IIf(IsNull(Trim(rs1!ADDRESS1)), " ", rs1!ADDRESS1 & "  ")
                Print #1, Tab(20 + 15); IIf(IsNull(Trim(rs1!ADDRESS2)), " ", rs1!ADDRESS2 & "  ")
                Print #1, Tab(20 + 15); IIf(IsNull(Trim(rs1!ADDRESS3)), " ", rs1!ADDRESS3 & "  ")
                Print #1, ""
            Else
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                
            End If
            Print #1, ""
            Print #1, ""; Chr(27) + Chr(15)
            Print #1, Tab(31); "RS. " & Trim(Right(toword(RS!baa), Len(toword(RS!baa)) - 6)) & ""
            Print #1, Tab(31); "( " & Format(Trim(Str(RS!baa)), "0.00") & " )"
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Close #1
            PrintOption.Show
      Else
            MsgBox "Bank Advice Amount 0 "
            Text1.SetFocus
      End If
 End If
End Function

Private Sub print1_Click()
    Dim frompage, topage, copies As Integer
    c1.Flags = 0
    c1.Max = 11
    c1.Flags = &H40000 Or &H4&
    c1.ShowPrinter
    frompage = c1.frompage
    topage = c1.topage
    copies = c1.copies
    If c1.Flags And &H20 Then
        c2.ShowSave
    End If
    
    End Sub
   

