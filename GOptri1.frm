VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form GOptrial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      Caption         =   "Dos Report"
      Height          =   390
      Left            =   4860
      TabIndex        =   16
      Top             =   2250
      Width           =   1590
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Crystal Report"
      Height          =   315
      Left            =   2970
      TabIndex        =   15
      Top             =   2310
      Width           =   1575
   End
   Begin VB.CommandButton print1 
      Height          =   345
      Left            =   2010
      Picture         =   "GOptri1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4710
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   1560
      Picture         =   "GOptri1.frx":0172
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5820
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2370
      TabIndex        =   11
      Text            =   "100 %"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   4095
      Left            =   1800
      TabIndex        =   10
      Top             =   2970
      Visible         =   0   'False
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7223
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      MaxLength       =   99999999
      RightMargin     =   20000
      TextRTF         =   $"GOptri1.frx":05C3
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
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   4200
      TabIndex        =   1
      Top             =   1410
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   2490
      TabIndex        =   0
      Top             =   1410
      Width           =   1545
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   5265
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      _Version        =   393216
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   5760
      TabIndex        =   2
      Top             =   765
      Visible         =   0   'False
      Width           =   345
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   5175
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      _Version        =   393216
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox alpha 
      Height          =   345
      Left            =   5625
      TabIndex        =   3
      Top             =   675
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   1
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   450
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Alphabat"
      Height          =   195
      Left            =   4635
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label3 
      Caption         =   "  As On :"
      Height          =   285
      Left            =   4740
      TabIndex        =   8
      Top             =   690
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From The Date"
      Height          =   195
      Left            =   4620
      TabIndex        =   7
      Top             =   750
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Gen. Ledger Desc."
      Height          =   195
      Left            =   4650
      TabIndex        =   6
      Top             =   750
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "GOptrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Dim rs As Recordset

Private Sub alpha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.Text) <> "" Then
                rs.Open "select * from gledger where " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not rs.BOF Then
            rs.Find "gledger='" + Trim(COMBOGENLEDGER.Text) + "'"
            If rs.EOF Then
                COMBOGENLEDGER.SetFocus
            End If
        Else
            COMBOGENLEDGER.SetFocus
        End If
        rs.close
    End If
End Sub

Private Sub Command1_Click()
        r1.Visible = False
        Me.print1.Visible = False
        Me.export.Visible = False
        Me.Combo1.Visible = False
        Me.Command1.Visible = False
End Sub

Private Sub Commandreturn_Click()
    Unload Me
End Sub
Private Sub Commandshow_Click()
Genrate
'PrintOption.Show

Unload PrintOption
Load PrintOption
'r1.LoadFile (App.Path + "\vipin.txt")
'r1.Visible = True
'Me.print1.Visible = True
'Me.export.Visible = True
'Me.Combo1.Visible = True
'Me.Command1.Visible = True
End Sub

Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub date1_LostFocus()
    If Trim(date1.Text) <> "" Then
        If Not checkdate(Trim(date1.Text), date1) Then
            date1.SetFocus
        End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub date2_LostFocus()
    If Trim(date2.Text) <> "" Then
        If Not checkdate(Trim(date2.Text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Option1.value = True
Me.r1.Top = 10
Me.r1.Left = 10
'''''Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
'''''    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
'''''        Unload VB.Screen.ActiveForm
'''''    End If
'''''Loop
Me.Top = 0
Me.Left = 0

'Set CON = New ADODB.Connection
'CON.CursorLocation = adUseClient
Set rs = New ADODB.Recordset

''    With CON
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
'        .Open
'    End With
        rs.Open "select * from gledger where " & stringyear, CON, adOpenDynamic, adLockReadOnly, adCmdText
    
    If Not rs.BOF Then
        Do While Not rs.EOF
            COMBOGENLEDGER.AddItem Trim(rs!gledger)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    

    rs.close
    'CNSetup
    rs.Open "Select * from setup where " & stringyear & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
    date1.Text = rs!yarfrom
    date2.Text = rs!yarto
   
    rs.close
End Sub

Private Sub print_Click()
Rsinvoicea.Open "select GenLedger,  SubLedger , sum(amount) as INVAmount from invoicea where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by invoicedate", CON, adOpenDynamic, adLockReadOnly, adCmdText
   
       RsCREDITa.Open "select * from CREDITa where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
       RsCnf1a.Open "select * from Cnf1a where  " & stringyear & " and pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
       Rsdnfa.Open "select * from dnfa where  " & stringyear & " and pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
       RsCnf1B.Open "select * from Cnf1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
       RsdnfB.Open "select * from dnfB where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
       RScasha.Open "select * from casha where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by invoicedate", CON, adOpenDynamic, adLockReadOnly, adCmdText
  
End Sub


'**********Gen opening Trial
Sub Genrate()
  Dim Diff1 As Double
    Dim Diff2 As Double
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30
    Set rs = New ADODB.Recordset
    main.reportname = "Gen. Ledger Trial"
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim xtemp As String
    Dim trs As ADODB.Recordset
    Dim FooterYes As Boolean
     FooterYes = False
    Set trs = New ADODB.Recordset
        paperWidth = 150
        T1 = 10
        T2 = 25
        T3 = 40
        T4 = 55
        T5 = 70
        T6 = 85
        T7 = 100
        T8 = 115
        MaxLine = 72
        Diff1 = 0
        Diff2 = 0
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Dim Pno As Integer
        Dim GopenBal As Double
        Dim GopenDr As Double
        Dim GopenCr As Double
        Dim GopenCl As Double
        
        Dim GSumDr  As Double
        Dim GSumCr  As Double
        GSumDr = 0
        GSumCr = 0
        
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        main.reportdata
        main.repors.Find "reportname='" + Trim(main.reportname) + "'"
        MaxLine = main.repors!totalline
        If main.repors!comp = True Then
            paperWidth = Int(main.repors!totalcolumn * 1.75)
        Else
            paperWidth = main.repors!totalcolumn
        End If
        Open "" + App.Path + "\vipin.txt" For Output As #1
        Line = 0
        Pno = 1
header:
        Dim I As Integer
        For I = 1 To main.repors!TopMargin
            Print #1, ""
            Line = Line + 1
        Next
        If FooterYes = True Then
            Do While Line <= 72
                    Print #1, " "
                    Line = Line + 1
            Loop
            Line = 0
            FooterYes = False
        End If
        If kkk.State = 1 Then kkk.close
        CNSetup
        kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, ""
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab(115); "Page No:  " & Pno
            Print #1, Tab((((paperWidth - (Len(Trim(kkk!CNAME)) * 2)) / 2) - 15) + LEFTM); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!CNAME))
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1)); Chr(27) + Chr(14)
            Line = Line + 5
        End If
        If trs.State = 1 Then
            trs.close
        End If
        trs.Open "select * from treport where  " & stringyear & " and userid=" & main.UId & " order by sno", CON, adOpenDynamic, adLockReadOnly, adCmdText
        Dim rs7 As New ADODB.Recordset
        rs7.Open "Select * from setup where " & stringyear & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
        Print #1, ""
        Print #1, Tab(((paperWidth - (Len(Trim("Gen.Opening Trial Balance As On April ` " & Right(rs7!yarfrom, 4))))) / 2) + LEFTM); "Gen.Opening Trial Balance As On April ` " & Right(rs7!yarfrom, 4)
        Print #1, ""
        Print #1, Chr(27) + Chr(71); Tab(LEFTM); repli("-", paperWidth - 10)
        Print #1, Tab(8); "Gen. Ledger Description"; Tab(67); "Amount (Dr.)"; Tab(96); "Amount (Cr.)"
        Print #1, Tab(LEFTM); repli("-", paperWidth - 10); Chr(27) + Chr(72);
        Print #1, ""
        Line = Line + 7
        trs.close
        If called1 Then
           called1 = False
           GoTo printagain1
        End If
        If rs.State = 1 Then rs.close
        Dim DbB As Double
        Dim CrB As Double
        DbB = 0
        DrB = 0
        If rs.State = 1 Then rs.close
         rs.Open "select  * from gledger  where  " & stringyear & " and yearopening <>0 order by gledger ", CON, adOpenDynamic, adLockReadOnly, adCmdText
        
        While Not rs.EOF
              DbB = 0
              CrB = 0
              DbB = IIf(rs!YEAROPENING > 0, rs!YEAROPENING, 0)
              CrB = IIf(rs!YEAROPENING < 0, rs!YEAROPENING, 0)
              Print #1, Tab(1); rs!gledger; Tab(65); IIf(DbB <> 0, rsets(Trim(Format(DbB, "0.00")), 12), ""); Tab(96); IIf(CrB <> 0, rsets(Trim(Format(Str(Abs(CrB)), "0.00")), 12), "")
              Line = Line + 1
              If Line > MaxLine - 9 Then
                        called1 = True
                        FooterYes = True
                        Pno = Pno + 1
                        GoTo header
printagain1:
                        called1 = False
                End If
                If Not rs.EOF Then
                    rs.MoveNext
                End If
                GSumDr = GSumDr + DbB
                GSumCr = GSumCr + Abs(CrB)
            Wend
            
printfooter:
            If GSumDr > GSumCr Then
               Diff1 = GSumDr - GSumCr
            Else
               Diff2 = GSumCr - GSumDr
            End If
            Print #1, Tab(LEFTM); repli("-", paperWidth - 10)
            Print #1, Tab(LEFTM); "* * * NET TOTAL * * * "; Tab(65); IIf(GSumDr <> 0, rsets(Format(Trim(GSumDr), "0.00"), 12), ""); Tab(96); IIf(GSumCr <> 0, rsets(Format(Trim(GSumCr), "0.00"), 12), "")
            Print #1, Tab(LEFTM); repli("-", paperWidth - 10)
            Line = Line + 3
            Do While Line <= 72
                Print #1, " "
                Line = Line + 1
            Loop
            If trs.State = 1 Then trs.close
            Close #1
End Sub
Private Sub print1_Click()
     c1.PrinterDefault = True
    c1.ShowPrinter
    printnow
End Sub
