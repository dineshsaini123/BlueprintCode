VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPartyWiseAgentwise 
   Caption         =   "Party Wise & Executive Wise Gross and Net Sale"
   ClientHeight    =   4992
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7008
   Icon            =   "frmPartyWiseAgentwise.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4992
   ScaleWidth      =   7008
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1_excel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export To Excel"
      Height          =   645
      Left            =   2835
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3510
      Width           =   1185
   End
   Begin VB.ComboBox cboSt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmPartyWiseAgentwise.frx":000C
      Left            =   1545
      List            =   "frmPartyWiseAgentwise.frx":000E
      TabIndex        =   3
      Top             =   1575
      Width           =   4935
   End
   Begin VB.CheckBox Check1_bill 
      Caption         =   "Including School && Bill No"
      Height          =   285
      Left            =   1590
      TabIndex        =   12
      Top             =   2865
      Width           =   2355
   End
   Begin VB.CheckBox Check1_school 
      Caption         =   "Including School"
      Height          =   240
      Left            =   1590
      TabIndex        =   11
      Top             =   2565
      Width           =   2370
   End
   Begin VB.TextBox txtParty1 
      Height          =   315
      Left            =   5505
      TabIndex        =   10
      Top             =   300
      Visible         =   0   'False
      Width           =   975
   End
   Begin Crystal.CrystalReport cr 
      Left            =   780
      Top             =   3060
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CommandPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   645
      Left            =   1560
      Picture         =   "frmPartyWiseAgentwise.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3510
      Width           =   1185
   End
   Begin VB.CommandButton CommandReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   645
      Left            =   4095
      Picture         =   "frmPartyWiseAgentwise.frx":0BF4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3510
      Width           =   1185
   End
   Begin VB.ComboBox cboagent 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmPartyWiseAgentwise.frx":17D8
      Left            =   1545
      List            =   "frmPartyWiseAgentwise.frx":1809
      TabIndex        =   2
      Top             =   1155
      Width           =   4935
   End
   Begin VB.TextBox txtParty 
      Height          =   315
      Left            =   1545
      TabIndex        =   1
      Top             =   780
      Width           =   4875
   End
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   315
      Left            =   1545
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txtto 
      Height          =   315
      Left            =   3405
      TabIndex        =   4
      Top             =   240
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   42409
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "St. Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   13
      Top             =   1590
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Executive Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   7
      Top             =   1170
      Width           =   1650
   End
   Begin VB.Label lblAson 
      Caption         =   "To"
      Height          =   255
      Left            =   3045
      TabIndex        =   6
      Top             =   240
      Width           =   315
   End
   Begin VB.Label lblParty 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   840
      Width           =   1275
   End
End
Attribute VB_Name = "frmPartyWiseAgentwise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_excel_Click()

Dim str1 As String
  
If txtParty.Text = "" Then
   txtParty1.Text = ""
End If


str1 = ""

 If txtParty.Text <> "" Then
    str1 = "SUBLEDGER = '" & txtParty1.Text & "'"
 End If

If cboSt.Text <> "" Then
      If str1 = "" Then
         str1 = "states = '" & cboSt.Text & "'"
      Else
         str1 = str1 & " and states = '" & cboSt.Text & "'"
      End If
End If


If cboagent.Text <> "" Then
      If str1 = "" Then
         str1 = "agentname = '" & cboagent.Text & "'"
      Else
         str1 = str1 & " and agentname = '" & cboagent.Text & "'"
      End If
End If

Dim str_date As String
str_date = "(invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & txtto.value & "',103))"

If str1 = "" Then
   str1 = str_date
Else
   str1 = str1 & " and  " & str_date
End If

  
  
'str_1 = "select  Code,SUBLEDGER,states,agentname,ScName,INVOICENO,BOOKNAME,QUANTITY,Gross,Net from Repwise_party_billwisegrossale where " & str1 & " order by INVOICENO "

If Check1_bill.value = 1 Then
   createExcel "select  Code,SUBLEDGER,states,agentname,ScName,INVOICENO,BOOKNAME,QUANTITY,Gross,Net from Repwise_party_billwisegrossale where " & str1 & " order by INVOICENO ", "1"
ElseIf Check1_school.value = 1 Then
   createExcel "SELECT distinct Code,SUBLEDGER,states,agentname,ScName,Gross,Net from Repwise_party_billwisegrossale where " & str1 & "", "2"
End If


End Sub
Private Sub CommandPrint_Click()
   
  Dim str1 As String
  
  DSNNew
  
  If txtParty.Text = "" Then
     txtParty1.Text = ""
  End If
  
  
  str1 = ""
  
   If txtParty.Text <> "" Then
      str1 = "{SaleRegisterAgentwise.subledger} = '" & txtParty1.Text & "'"
   End If
  
  If cboSt.Text <> "" Then
        If str1 = "" Then
           str1 = "{SaleRegisterAgentwise.states} = '" & cboSt.Text & "'"
        Else
           str1 = str1 & " and {SaleRegisterAgentwise.states} = '" & cboSt.Text & "'"
        End If
  End If
  
  
  If cboagent.Text <> "" Then
        If str1 = "" Then
           str1 = "{SaleRegisterAgentwise.agentname} = '" & cboagent.Text & "'"
        Else
           str1 = str1 & " and {SaleRegisterAgentwise.agentname} = '" & cboagent.Text & "'"
        End If
  End If
  
  
    ''If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then

    MainMenu.cr1.Reset
    If Check1_school.value = 1 Then
       MainMenu.cr1.ReportFileName = rptPath & "/Agentsaleregister_gsale_withSchool.rpt"
    ElseIf Check1_bill.value = 1 Then
       MainMenu.cr1.ReportFileName = rptPath & "/Agentsaleregister_gsale_withSchool_billwise.rpt"
    Else
       MainMenu.cr1.ReportFileName = rptPath & "/Agentsaleregister_gsale.rpt"
    End If
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If str1 = "" Then
    MainMenu.cr1.ReplaceSelectionFormula "({SaleRegisterAgentwise.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {SaleRegisterAgentwise.invoicedate}<=datevalue('" & Format(txtto.value, "MM/dd/yyyy") & "')) "
    Else
    MainMenu.cr1.ReplaceSelectionFormula "({SaleRegisterAgentwise.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {SaleRegisterAgentwise.invoicedate}<=datevalue('" & Format(txtto.value, "MM/dd/yyyy") & "')) and " & str1
    End If
    MainMenu.cr1.Formulas(0) = "fdate='" & txtFrom.value & "'"
    MainMenu.cr1.Formulas(1) = "tdate='" & txtto.value & "'"
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.Action = 1
 
 
 
 


End Sub

Private Sub CommandReturn_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.Top = 500
Me.Left = 500


   
    If RS.State = 1 Then RS.close
    RS.Open "select Rep as Representative from SalesRepQry order by Rep", CON_blue
    Me.cboagent.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cboagent.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
    RS.close


    If RS.State = 1 Then RS.close
    RS.Open "select distinct(states) as State from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & "", con
    While RS.EOF = False
       Me.cboSt.AddItem RS(0)
       RS.MoveNext
    Wend


txtFrom.value = from_date
txtto.value = to_date

End Sub

Private Sub txtParty_GotFocus()

If PopUpValue1 <> "" Then
   txtParty.Text = PopUpValue1
   txtParty1.Text = PopUpValue3
   
   PopUpValue1 = ""
   PopUpValue3 = ""
   
End If

End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    searchType = "party"
    value = "select distinct(Party),Code,subledger from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & "  order by party"
    popuplist_client value, con
    set_focus = True
End If

End Sub
