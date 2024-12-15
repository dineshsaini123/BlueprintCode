VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form bankadvice 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   1410
   ClientWidth     =   9195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "badvice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtadno 
      Height          =   345
      Left            =   3930
      TabIndex        =   3
      Top             =   480
      Width           =   1845
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   4185
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   7382
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   20000
      TextRTF         =   $"badvice.frx":000C
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
      Picture         =   "badvice.frx":008C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "badvice.frx":01FE
      Left            =   3900
      List            =   "badvice.frx":0211
      TabIndex        =   5
      Text            =   "100 %"
      Top             =   4710
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   2760
      Picture         =   "badvice.frx":0237
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   3330
      TabIndex        =   1
      Top             =   2160
      Width           =   1605
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   5100
      TabIndex        =   2
      Top             =   2160
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   3930
      TabIndex        =   0
      Top             =   1020
      Width           =   1845
   End
   Begin VB.Label Label2 
      Caption         =   "Advice No.(If do you want to start own advice no the use this box)"
      Height          =   435
      Left            =   570
      TabIndex        =   9
      Top             =   480
      Width           =   3075
   End
   Begin VB.Label Label1 
      Caption         =   "Invoice No."
      Height          =   285
      Left            =   1950
      TabIndex        =   7
      Top             =   1050
      Width           =   1305
   End
End
Attribute VB_Name = "bankadvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLSTRING As String
Dim rs As ADODB.Recordset

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
MainMenu.Toolbar1.Visible = True
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
    For I = 1 To (Printdlg.UpDown1.Value)
        X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(p.Port))
    Next
    Printdlg.UpDown1.Value = 1
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
  Set rs = New ADODB.Recordset
  Dim rs1 As New ADODB.Recordset
  rs1.Open "Select * from setup where " & stridnyear & "", CON, adOpenKeyset, adLockOptimistic, adCmdText
  If rs1.RecordCount > 0 Then
     txtadno.Text = IIf(IsNull(rs1!bankadviceno), 0, (rs1!bankadviceno))
  End If
  

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
   Set rs = New ADODB.Recordset
   rs.Open "select * from invoicea where  " & stridnyear & " and invoiceno=" + Text1.Text + "", CON, adOpenKeyset, adLockReadOnly, adCmdText
   If rs.RecordCount > 0 Then
        If rs!baa > 0 Then
            rs1.Open "Select * from setup where " & stridnyear & "", CON, adOpenKeyset, adLockOptimistic, adCmdText
            rs1.MoveFirst
            If rs.BOF Then
                MsgBox "Invoice not found"
               Exit Function
            End If
            Open "" + App.Path + "\vipin.txt" For Output As #1
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            'Print #1, Tab(16); (rs1!bankadviceno + 1); Tab(66); rs!invoicedate
            Print #1, Tab(16); txtadno.Text; Tab(66); rs!invoicedate
            rs1!bankadviceno = txtadno.Text + 1 'rs1!bankadviceno + 1
            txtadno.Text = txtadno + 1
            rs1.Update
            Print #1, ""
            Print #1, Tab(20); Trim(rs!through);
            If Trim(rs!through) <> "" Then
               Print #1, " ,"
            Else
               Print #1, ""
            End If
            Print #1, Tab(20); Trim(rs!through1)
            Print #1, ""
            Print #1, ""
            Print #1, Tab(20); Trim(rs!biltyno); Tab(36); IIf(IsNull(rs!BILTYDATE), " / /   ", rs!BILTYDATE)
            Print #1, Tab(20); Trim(rs!INVOICENO); Tab(36); IIf(IsNull(rs!invoicedate), " / /   ", rs!invoicedate)
            If rs1.State = 1 Then rs1.Close
            rs1.Open "select * from sledger where  " & stridnyear & " and gledger='SUNDRY DEBTORS' and subledger='" + Trim(rs!subledger) + "'", CON, adOpenKeyset, adLockReadOnly
            Print #1, ""
            Print #1, ""
            If Not rs1.BOF Then
                Print #1, Tab(32); IIf(IsNull(rs1!DESCFORINVOICE), "", rs1!DESCFORINVOICE)
                Print #1, Tab(32); IIf(IsNull(Trim(rs1!address1)), " ", rs1!address1 & "  ") ' IIf(IsNull(Trim(rs1!ADDRESS2)), "", rs1!ADDRESS2)
                Print #1, ""
               ' Print #1, Tab(32); IIf(IsNull(Trim(rs1!ADDRESS3)), "", rs1!ADDRESS3)
            Else
                Print #1, ""
                Print #1, ""
                Print #1, ""
            End If
            Print #1, Tab(20); Right(toword(rs!baa), Len(toword(rs!baa)) - 6)
            Print #1, Tab(20); "( " & Format(Trim(Str(rs!baa)), "0.00") & " )"
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

