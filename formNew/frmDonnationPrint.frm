VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmDonnationPrint 
   Caption         =   "Print "
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6948
   Icon            =   "frmDonnationPrint.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3480
   ScaleWidth      =   6948
   Begin VB.CheckBox Check1_bkwise 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Book Wise Donation"
      Height          =   315
      Left            =   4020
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   780
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cbostate 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmDonnationPrint.frx":000C
      Left            =   1560
      List            =   "frmDonnationPrint.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   780
      Width           =   2400
   End
   Begin VB.TextBox txtScid 
      Height          =   315
      Left            =   6120
      TabIndex        =   7
      Top             =   1800
      Width           =   315
   End
   Begin VB.ComboBox cbotrpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      ItemData        =   "frmDonnationPrint.frx":003C
      Left            =   1560
      List            =   "frmDonnationPrint.frx":004F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   4290
   End
   Begin VB.TextBox txtParty 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   4515
   End
   Begin VB.ComboBox cboPayment 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmDonnationPrint.frx":00D1
      Left            =   1560
      List            =   "frmDonnationPrint.frx":00E7
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1140
      Width           =   2400
   End
   Begin Crystal.CrystalReport cr 
      Left            =   6360
      Top             =   180
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint_7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   720
      Left            =   1560
      Picture         =   "frmDonnationPrint.frx":012B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label2_paymode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "State :"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label2_paymode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1740
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2775
      Left            =   60
      Top             =   660
      Width           =   6495
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type Of Reports :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   300
      Width           =   1530
   End
   Begin VB.Label Label1_ff2 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search"
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label Label2_paymode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mode :"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1275
   End
End
Attribute VB_Name = "frmDonnationPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbostate_Click()

If cbotrpt.Text = "School Wise Donation" Then
   If cbostate <> "" Then
   Screen.MousePointer = vbHourglass
    con.Execute "exec update_SchoolList '" & cbostate & "'"
   Screen.MousePointer = vbDefault
    End If
End If

End Sub

Private Sub cbotrpt_Click()

    Label2_paymode(2).Visible = False
    cbostate.Visible = False
    txtParty.Visible = False
    Label1_ff2.Visible = False
    Check1_bkwise.Visible = False


If (cbotrpt.Text = "State Wise") Then
   Label2_paymode(0).Visible = False
   cboPayment.Visible = False
   
   txtParty.Visible = False
   txtScId.Visible = False
   Label1_ff2.Visible = False
   
   cbostate.Visible = True
   Label2_paymode(2).Visible = True
   Check1_bkwise.Visible = True
 
ElseIf cbotrpt = "State Wise & Rep. Wise Donation" Then
   
    Label2_paymode(2).Visible = True
    cbostate.Visible = True
    txtParty.Visible = True
    Label1_ff2.Visible = True
   
ElseIf cbotrpt.Text = "School Wise Donation" Then
   Label2_paymode(0).Visible = False
   cboPayment.Visible = False
   
   Label2_paymode(1).Caption = "School :"
   txtParty.Visible = True
   txtScId.Visible = True
   
   cbostate.Visible = True
   Label2_paymode(2).Visible = True
   
 
   
ElseIf cbotrpt.Text = "Rep. Wise Donation" Then
   Label2_paymode(0).Visible = True
   cboPayment.Visible = True
   Label2_paymode(1).Caption = "Rep. Name :"
   txtParty.Visible = True
   txtScId.Visible = True
   
   cbostate.Visible = False
   Label2_paymode(2).Visible = False
   Label1_ff2.Visible = True
   
   
End If
End Sub

Private Sub cmdPrint_7_Click()

DSNNew

txtScId = ""
'====================================
If cboPayment.Text <> "" Then
    Dim st_ As String
    If rs1.State = 1 Then rs1.close
    rs1.Open "select DNo from DonnationMain where PaymentMode='" & cboPayment.Text & "' and party_  is null", con
    While rs1.EOF = False
        st_ = ""
        If RS.State = 1 Then RS.close
        RS.Open "select subledger from DonnationQry_ where dno=" & rs1!dno & " and subledger is not null", con, adOpenDynamic, adLockReadOnly
        While RS.EOF = False
             If st_ = "" Then
                 st_ = RS!subledger & vbCrLf
             Else
                 st_ = st_ & ", " & RS!subledger
             End If
             RS.MoveNext
         Wend
         con.Execute "update DonnationMain set party_='" & st_ & "'  where DNo=" & rs1!dno & ""
         rs1.MoveNext
    Wend
    
    If MsgBox("Want to View ?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

End If

'====================================


If cbotrpt.Text = "State Wise" Then

    If cbostate <> "" Then
    con.Execute "exec update_SchoolList '" & cbostate & "'"
    Else
    con.Execute "exec update_SchoolList ''"
    End If
    
    
    If Check1_bkwise.value = 1 Then
        cr.Reset
        cr.ReportFileName = rptPath & "/ExtraDis_BkWise.rpt"
        cr.ReplaceSelectionFormula "{donation_Qry.states}='" & cbostate.Text & "'"
        cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
        cr.WindowShowPrintSetupBtn = True
        cr.WindowShowPrintBtn = True
        cr.WindowShowExportBtn = True
        cr.WindowState = crptMaximized
        cr.Action = 1
    Else
        
        cr.Reset
        cr.ReportFileName = rptPath & "/ExtraDis_St.rpt"
        If cbostate.Text <> "" Then
        cr.ReplaceSelectionFormula "{donation_StateWise.state}='" & cbostate.Text & "'"
        End If
        cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
        cr.WindowShowPrintSetupBtn = True
        cr.WindowShowPrintBtn = True
        cr.WindowShowExportBtn = True
        cr.WindowState = crptMaximized
        cr.Action = 1
        
    End If
    



ElseIf cbotrpt.Text = "Rep. Wise Donation" Then
  

   
   st11 = ""
   If (cboPayment.Text <> "") Then
        If st11 = "" Then
        If cboPayment.Text = "(-) Round Of Amt" Then
           st11 = "{DonnationMain.RoundOfAAmt}<0"
        
        ElseIf cboPayment.Text = "Advance Amt" Then
           st11 = "{DonnationMain.AdvAmt}>0"
        Else
           st11 = "{DonnationMain.PaymentMode}='" & cboPayment.Text & "'"
        End If
        End If
   End If


   
   If (txtParty.Text <> "") Then
        If st11 = "" Then
            st11 = "{DonnationMain.RepName}='" & txtParty & "'"
        Else
            st11 = st11 & " and " & "{DonnationMain.RepName}='" & txtParty & "'"
        End If
   End If


    cr.Reset
    cr.ReportFileName = rptPath & "/Donnation_Rep.rpt"
    If st11 <> "" Then
       cr.ReplaceSelectionFormula st11
    End If
    
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If txtParty = "" Then
    cr.Formulas(0) = "sales_salesret='" & "DONATION DETAILS" & "'"
    Else
    cr.Formulas(0) = "sales_salesret='" & "DONATION DETAILS OF MR " + txtParty & "'"
    End If
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowPrintBtn = True
    cr.WindowShowExportBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1

ElseIf cbotrpt.Text = "State Wise & Rep. Wise Donation" Then
   
   st11 = ""

   If (cboPayment.Text <> "") Then
        If st11 = "" Then
          st11 = "{DonnationMain.PaymentMode}='" & cboPayment.Text & "'"
        End If
   End If


   
   If (txtParty.Text <> "") Then
        If st11 = "" Then
            st11 = "{DonnationMain.RepName}='" & txtParty & "'"
        Else
            st11 = st11 & " and " & "{DonnationMain.RepName}='" & txtParty & "'"
        End If
   End If


    cr.Reset
    cr.ReportFileName = rptPath & "/Donnation_RepSt.rpt"
    If st11 <> "" Then
       cr.ReplaceSelectionFormula st11
    End If
    
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.Formulas(0) = "sales_salesret='" & "DONATION DETAILS OF MR " + txtParty & "'"
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowPrintBtn = True
    cr.WindowShowExportBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
    
ElseIf cbotrpt.Text = "Party & Title Code Wise Adj. Discount" Then

    con.Execute "UPDATE a SET a.remarks_ = b.Remarks  FROM SalesAdjustmentDet AS a INNER JOIN SalesAdjustment AS b ON (a.dno = b.dno) "
    DoEvents
    DoEvents
    DoEvents

    cr.Reset
    cr.ReportFileName = rptPath & "/Party&TCodeAdjDis.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowPrintBtn = True
    cr.WindowShowExportBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1


ElseIf cbotrpt.Text = "School Wise Donation" Then

   'Dim st11 As String
   
   st11 = ""
   

    If (cbostate <> "" Or cbostate <> ".") Then
       con.Execute "exec update_SchoolList '" & cbostate & "'"
    End If

   If (cbostate.Text <> "") Then
        If st11 = "" Then
          st11 = "{donationQry.State}='" & cbostate.Text & "'"
        Else
          st11 = st11 & " and " & "{donationQry.State}='" & cbostate & "'"
        End If
   End If


   
   If (txtParty.Text <> "") Then
   
    If st11 = "" Then
      st11 = "{donationQry.ScName}='" & txtParty & "'"
    Else
      st11 = st11 & " and " & "{donationQry.ScName}='" & txtParty & "'"
    End If
      
   End If


    If txtParty <> "" Then
       
     If txtScId <> "" Then
       If RS.State = 1 Then RS.close
       RS.Open "select top 1 [state] from collegeView_ind where CollegeID='" & txtScId & "'", CON_blue
       If RS.EOF = False Then
          con.Execute "exec update_SchoolList '" & RS(0) & "'"
       End If
     End If
     
    End If



    cr.Reset
    cr.ReportFileName = rptPath & "/Donnation_shoolWise.rpt"
    
    If st11 <> "" Then
    cr.ReplaceSelectionFormula st11
    End If
    
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.Formulas(0) = "sales_salesret='" & "DONATION DETAILS OF MR " + cmbAgentName & "'"
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowPrintBtn = True
    cr.WindowShowExportBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1

End If

txtParty = ""



End Sub
Private Sub Form_GotFocus()
If PopUpValue1 <> "" Then
   txtParty = PopUpValue1
   PopUpValue1 = ""
End If
End Sub

Private Sub Form_Load()

If RS.State = 1 Then RS.close
RS.Open "select distinct States from SLEDGER order by States", con
cbostate.Clear
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cbostate.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If



cbotrpt.ListIndex = 0
cbostate.ListIndex = 0

Me.Width = 6800
Me.Height = 5000

BackColorFrom Me
End Sub
Private Sub txtParty_GotFocus()

If (cbotrpt.Text = "State Wise" Or cbotrpt.Text = "Rep. Wise Donation" Or cbotrpt.Text = "State Wise & Rep. Wise Donation") Then
    If PopUpValue1 <> "" Then
       txtParty = PopUpValue1
       PopUpValue1 = ""
    End If
ElseIf cbotrpt.Text = "School Wise Donation" Then
    If PopUpValue1 <> "" Then
       txtParty = PopUpValue1
       txtScId = PopUpValue2
       
       PopUpValue1 = ""
       PopUpValue2 = ""
    End If
End If

End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   
   If cbotrpt.Text = "State Wise" Then
      searchType = "party"
      value = "select States from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & " group by States"
      popuplist_client value, con
      set_focus = True
   ElseIf (cbotrpt.Text = "Rep. Wise Donation" Or cbotrpt.Text = "State Wise & Rep. Wise Donation") Then
      searchType = "party"
      value = "select Rep as Representative from SalesRepQry order by Rep"
      popuplist_client value, CON_blue
      set_focus = True
   
   ElseIf cbotrpt.Text = "School Wise Donation" Then
   
        'If RS.State = 1 Then RS.close
        value = "SELECT  INVOICEA.ScName,INVOICEA.ScID,SchoolList.State FROM INVOICEA  INNER JOIN SchoolList ON dbo.INVOICEA.ScID = dbo.SchoolList.CollegeID where SchoolList.State='" & cbostate.Text & "'"
        
     
        searchType = "party"
        'value = "SELECT  INVOICEA.ScName,INVOICEA.ScID,SchoolList.State FROM INVOICEA INNER JOIN SchoolList ON dbo.INVOICEA.ScID = dbo.SchoolList.CollegeID where [SchoolList.State]='" & cbostate.Text & "'"
        popuplist_client value, con
        set_focus = True
        
        
        
   
   End If
   
 End If
   
End Sub
