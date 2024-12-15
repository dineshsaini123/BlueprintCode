VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEndPartTrans 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6384
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7392
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6384
   ScaleWidth      =   7392
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   144
      Top             =   4248
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "Refresh &Grid"
      Height          =   375
      Left            =   60
      TabIndex        =   14
      Top             =   5460
      Width           =   1515
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   6030
      TabIndex        =   6
      Top             =   5895
      Width           =   1110
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3588
      Left            =   48
      TabIndex        =   0
      Top             =   480
      Width           =   7260
      _cx             =   12806
      _cy             =   6329
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   12648447
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   -2147483647
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSMask.MaskEdBox T1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   6015
      TabIndex        =   2
      Top             =   4470
      Width           =   1155
      _ExtentX        =   2053
      _ExtentY        =   550
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox T1TEXT 
      Height          =   315
      Left            =   3285
      TabIndex        =   1
      Top             =   4470
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   550
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox T2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   6015
      TabIndex        =   4
      Top             =   4770
      Width           =   1155
      _ExtentX        =   2053
      _ExtentY        =   550
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox T2TEXT 
      Height          =   315
      Left            =   3285
      TabIndex        =   3
      Top             =   4800
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   550
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox T3TEXT 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   6015
      TabIndex        =   5
      Top             =   5130
      Width           =   1155
      _ExtentX        =   2053
      _ExtentY        =   550
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin VB.Label lblFrt 
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      Height          =   264
      Left            =   1692
      TabIndex        =   17
      Top             =   5544
      Visible         =   0   'False
      Width           =   876
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   312
      Index           =   21
      Left            =   36
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   4584
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   348
      Index           =   20
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   4584
   End
   Begin VB.Label T3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By BANK :"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3285
      TabIndex        =   13
      Top             =   5130
      Width           =   2700
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Net Amount"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3285
      TabIndex        =   12
      Top             =   4140
      Width           =   2700
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6015
      TabIndex        =   11
      Top             =   4140
      Width           =   1155
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Balance :"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3285
      TabIndex        =   10
      Top             =   5505
      Width           =   2700
   End
   Begin VB.Label labelbalance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   6012
      TabIndex        =   9
      Top             =   5496
      Width           =   1152
   End
   Begin VB.Label GROSS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6030
      TabIndex        =   8
      Top             =   45
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amount :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4635
      TabIndex        =   7
      Top             =   45
      Width           =   1275
   End
End
Attribute VB_Name = "frmEndPartTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k10 As Integer
Dim frt As String
Dim postage As String
Dim din_ As Integer
Dim rss1 As New ADODB.Recordset
Private Sub cmdref_Click()
             
If searchForm = "invoice" Then
  
  con.Execute "delete from  invoicectmp where invoiceno=" & invoice.I_NO.text & ""
  
  Set rss1 = New ADODB.Recordset

  rss1.Open "select * from invoicec where invoiceno='" & invoice.I_NO.text & "'", con, adOpenDynamic, adLockOptimistic
  If rss1.EOF = True Then

  
  Set RS = New ADODB.Recordset
  RS.Open "select * from INVOICEEND where " & stringyear & " and type='" & searchForm & "' ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
  If RS.EOF = False Then

    vs.rows = RS.RecordCount + 1
    For I = 1 To RS.RecordCount
    
      vs.TextMatrix(I, 0) = RS!text
      vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
      RS.MoveNext
    Next

  End If
  
  End If


ElseIf searchForm = "invoice_sp" Then
   
  Set rss1 = New ADODB.Recordset

  rss1.Open "select * from invoicec_sp where invoiceno='" & frmBookIssueSp.I_NO.text & "'", con, adOpenDynamic, adLockOptimistic
  If rss1.EOF = True Then
  
      Set RS = New ADODB.Recordset
      RS.Open "select * from INVOICEEND where " & stringyear & " and type='" & searchForm & "' ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
      If RS.EOF = False Then
    
        vs.rows = RS.RecordCount + 1
        For I = 1 To RS.RecordCount
        
          vs.TextMatrix(I, 0) = RS!text
          vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
          RS.MoveNext
        Next
    
      End If
  
  
  End If
  
  


End If

End Sub

Private Sub CommandReturn_Click()
k10 = 0

If searchForm = "invoice" Then

    'calc
    'invoice.labelbybank = T3TEXT
    frmEndPartTrans.Hide
    'invoice.Show
    invoice.Enabled = True
    invoice.Commandsave.Enabled = True

ElseIf searchForm = "invoice_sp" Then

    calc_invoice_sp
    frmEndPartTrans.Hide
    'invoice.Show
     frmBookIssueSp.Enabled = True
     frmBookIssueSp.Commandsave.Enabled = True

ElseIf searchForm = "invoice_spret" Then

    'calc_invoice_spRet
    'frmBookIssueSp_Ret.labelbybank = T3TEXT
    frmEndPartTrans.Hide
    'frmBookIssueSp_Ret.Show
    frmBookIssueSp_Ret.Enabled = True
    frmBookIssueSp_Ret.Commandsave.Enabled = True


ElseIf searchForm = "credititem" Then


    'calc_creditItem
    'Critnote.labelbybank = T3TEXT
    frmEndPartTrans.Hide
    Critnote.Enabled = True
    Critnote.Commandsave.Enabled = True

ElseIf searchForm = "cash" Then

    'calc_cash
    'countersale.labelbybank = T3TEXT
    frmEndPartTrans.Hide
    countersale.Enabled = True
    countersale.Commandsave.Enabled = True
    


ElseIf searchForm = "cashbasil" Then


    'calc_cashbasil
    'frmBasilSales.labelbybank = T3TEXT
    frmEndPartTrans.Hide
    'frmBasilSales.Show
    frmBasilSales.Enabled = True
    frmBasilSales.Commandsave.Enabled = True

ElseIf searchForm = "cashbasilret" Then


    'calc_cashbasilRet
    frmEndPartTrans.Hide
    'frmBasilSales_Ret.Show
    frmBasilSales_Ret.Enabled = True
    frmBasilSales_Ret.Commandsave.Enabled = True

ElseIf searchForm = "invoiceblue" Then

    calc_invblue
    frmEndPartTrans.Hide
    frmInvoice_blueprint.Enabled = True
    frmInvoice_blueprint.Commandsave.Enabled = True

End If


'frmEndPart.Hide



End Sub

Private Sub Form_Activate()

T1TEXT.Visible = True
T1.Visible = True
T2TEXT.Visible = True
T2.Visible = True
T3.Visible = True
T3TEXT.Visible = True
Label4.Visible = True
labelbalance.Visible = True



iniForm
T3TEXT.Enabled = True

If searchForm = "invoice" Then
    GROSS.Caption = invoice.totalamount - invoice.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    Me.T3 = "By Bank :"
ElseIf searchForm = "invoiceblue" Then
    GROSS.Caption = frmInvoice_blueprint.totalamount - frmInvoice_blueprint.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    Me.T3 = "By Bank :"

ElseIf searchForm = "cashbasil" Then
    GROSS.Caption = frmBasilSales.totalamount - frmBasilSales.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    Me.T3 = "By Cash Recd. :"
ElseIf searchForm = "cashbasilret" Then
    GROSS.Caption = frmBasilSales_Ret.totalamount - frmBasilSales_Ret.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    Me.T3 = "By Cash Recd. :"
ElseIf searchForm = "invoice_sp" Then
    GROSS.Caption = frmBookIssueSp.totalamount - frmBookIssueSp.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    Me.T3 = "By Bank :"

ElseIf searchForm = "invoice_spret" Then
    GROSS.Caption = frmBookIssueSp_Ret.totalamount - frmBookIssueSp_Ret.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    Me.T3 = "By Bank :"

ElseIf searchForm = "credititem" Then
    GROSS.Caption = Critnote.totalamount - Critnote.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    T3TEXT.Enabled = False

ElseIf searchForm = "cash" Then
    GROSS.Caption = countersale.totalamount - countersale.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    Me.T3 = "By Cash Recd. :"
    
ElseIf searchForm = "cashbasil" Then
    GROSS.Caption = countersale.totalamount - countersale.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    Me.T3 = "By Cash Recd. :"

ElseIf searchForm = "cashbasilret" Then
    GROSS.Caption = countersale.totalamount - countersale.totaldiscount
    GROSS.Caption = Format(GROSS.Caption, "0.00")
    'Me.T3 = "By Cash Recd. :"

End If


din_ = 0

If searchForm = "invoice" Then

  frt = "Yes"
  

  
  'If rs1.State = 1 Then rs1.close
  'rs1.Open "select freight from SLEDGER  where subledger='" & invoice.textbox.text & "'", con
  'If rs1.EOF = False Then
  
     'If rs1!freight = "No" Then
     If invoice.lblPartyfrt.Caption = "No" Then
        frt = "No"
        Label1(20).Visible = True
        Label1(21).Visible = True
        din_ = 1
        Timer1.Enabled = True
     Else
        Label1(20).Visible = False
        Label1(21).Visible = False
        Timer1.Enabled = False
        
     End If
     
  'End If
  
  
  

  postage = "Yes"
  
  If rs1.State = 1 Then rs1.close
  rs1.Open "select postage from SLEDGER  where subledger='" & invoice.textbox.text & "'", con
  If rs1.EOF = False Then
    
    
     
     If rs1!postage = "No" Then
        din_ = 2
        postage = "No"
        Label1(20).Visible = True
        Label1(21).Visible = True

        Timer1.Enabled = True
'     Else
'        Label1(20).Visible = False
'        Label1(21).Visible = False
'        Timer1.Enabled = False
        
     End If
     
    End If
     

  

End If


vs.SetFocus
vs.Col = 2


 
  

End Sub
Sub calc()



Dim I As Integer
invoice.otherdiscount = 0
invoice.otheramount = 0
    
    
For I = 1 To vs.rows - 1
        
    If RS.State = 1 Then
        RS.close
    End If
    
    If invoice.edit Then
       RS.Open "select * from invoicectmp where invoiceno=" & invoice.I_NO.text & " and " & stringyear & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
    Else
       RS.Open "select * from invoiceend where " & stringyear & " and type='" & searchForm & "' order by printorder", con, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    
    RS.Find "TEXT='" + Trim(vs.TextMatrix(I, 0)) + "'"
    If Not RS.EOF Then
        If Trim(RS!DebitorCredit) = Trim("Debit") Then
            If RS!rate > 0 Then
                vs.Col = 1
                invoice.otherdiscount = invoice.otherdiscount + ((invoice.totalamount - invoice.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                invoice.otherdiscount = invoice.otherdiscount - Val(vs.TextMatrix(I, 2))
            End If
        Else
            If RS!rate > 0 Then
                vs.Col = 1
                invoice.otheramount = invoice.otheramount + ((invoice.totalamount + invoice.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                invoice.otheramount = invoice.otheramount + Val(vs.TextMatrix(I, 2))
            End If
        End If
        End If
    
    'RS.close
    Next
    
    
invoice.mna.Caption = Format(myround(invoice.totalamount + invoice.otheramount - invoice.totaldiscount + invoice.otherdiscount, 2), "0.00")
Me.Label3 = Format(invoice.totalamount + invoice.otheramount - invoice.totaldiscount + invoice.otherdiscount, "0.00")
Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
invoice.labelbybank = Me.T3TEXT
  
  
 
End Sub
Sub calc_invblue()



Dim I As Integer
frmInvoice_blueprint.otherdiscount = 0
frmInvoice_blueprint.otheramount = 0
    
    
For I = 1 To vs.rows - 1
        
    If RS.State = 1 Then
        RS.close
    End If
    
    If frmInvoice_blueprint.edit Then
       RS.Open "select * from invoicectmp_blue where invoiceno=" & frmInvoice_blueprint.I_NO.text & " and " & stringyear & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
    Else
       RS.Open "select * from invoiceend where " & stringyear & " and type='" & searchForm & "' order by printorder", con, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    
    RS.Find "TEXT='" + Trim(vs.TextMatrix(I, 0)) + "'"
    If Not RS.EOF Then
        If Trim(RS!DebitorCredit) = Trim("Debit") Then
            If RS!rate > 0 Then
                vs.Col = 1
                frmInvoice_blueprint.otherdiscount = frmInvoice_blueprint.otherdiscount + ((frmInvoice_blueprint.totalamount - frmInvoice_blueprint.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                frmInvoice_blueprint.otherdiscount = frmInvoice_blueprint.otherdiscount - Val(vs.TextMatrix(I, 2))
            End If
        Else
            If RS!rate > 0 Then
                vs.Col = 1
                frmInvoice_blueprint.otheramount = frmInvoice_blueprint.otheramount + ((frmInvoice_blueprint.totalamount + frmInvoice_blueprint.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                frmInvoice_blueprint.otheramount = frmInvoice_blueprint.otheramount + Val(vs.TextMatrix(I, 2))
            End If
        End If
        End If
    
    'RS.close
    Next
    
    
frmInvoice_blueprint.mna.Caption = Format(myround(frmInvoice_blueprint.totalamount + frmInvoice_blueprint.otheramount - frmInvoice_blueprint.totaldiscount + frmInvoice_blueprint.otherdiscount, 2), "0.00")
Me.Label3 = Format(frmInvoice_blueprint.totalamount + frmInvoice_blueprint.otheramount - frmInvoice_blueprint.totaldiscount + frmInvoice_blueprint.otherdiscount, "0.00")
Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
frmInvoice_blueprint.labelbybank = Me.T3TEXT
  
  
 
End Sub

Sub calc_invoice_sp()



Dim I As Integer
frmBookIssueSp.otherdiscount = 0
frmBookIssueSp.otheramount = 0
    
    
For I = 1 To vs.rows - 1
        
    If RS.State = 1 Then
        RS.close
    End If
    
    If frmBookIssueSp.edit Then
       RS.Open "select * from invoicec_Sp  where invoiceno=" & frmBookIssueSp.I_NO.text & " and " & stringyear & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
    Else
       RS.Open "select * from invoiceend where " & stringyear & " and type='" & searchForm & "' order by printorder", con, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    
    RS.Find "TEXT='" + Trim(vs.TextMatrix(I, 0)) + "'"
    If Not RS.EOF Then
        If Trim(RS!DebitorCredit) = Trim("Credit") Then
            If RS!rate > 0 Then
                vs.Col = 1
                'frmBookIssueSp.otherdiscount = frmBookIssueSp.otherdiscount + ((frmBookIssueSp.totalamount - frmBookIssueSp.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                'frmBookIssueSp.otherdiscount = frmBookIssueSp.otherdiscount + Val(vs.TextMatrix(I, 2))
                frmBookIssueSp.otheramount = frmBookIssueSp.otheramount + Val(vs.TextMatrix(I, 2))
            End If
        Else
            If RS!rate > 0 Then
                'vs.Col = 1
                'frmBookIssueSp.otheramount = frmBookIssueSp.otheramount - ((frmBookIssueSp.totalamount - frmBookIssueSp.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                
                frmBookIssueSp.otheramount = frmBookIssueSp.otheramount - Val(vs.TextMatrix(I, 2))
            End If
        End If
        End If
    
    'RS.close
    Next
    
frmBookIssueSp.mna.Caption = Format(myround(frmBookIssueSp.totalamount + frmBookIssueSp.otheramount - frmBookIssueSp.totaldiscount - frmBookIssueSp.otherdiscount, 2), "0.00")
frmBookIssueSp.mna.Caption = Format(myround(frmBookIssueSp.totalamount + frmBookIssueSp.otheramount - frmBookIssueSp.totaldiscount - frmBookIssueSp.otherdiscount, 2), "0.00")
Me.Label3 = Format(frmBookIssueSp.totalamount + frmBookIssueSp.otheramount - frmBookIssueSp.totaldiscount - frmBookIssueSp.otherdiscount, "0.00")
Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
  
  
  
 
End Sub
Sub calc_invoice_spRet()



Dim I As Integer

frmBookIssueSp_Ret.otherdiscount = 0
frmBookIssueSp_Ret.otheramount = 0
     
    
For I = 1 To vs.rows - 1
        
    If RS.State = 1 Then
        RS.close
    End If
    
    If frmBookIssueSp_Ret.edit Then
       RS.Open "select * from invoicec_Spret  where invoiceno=" & frmBookIssueSp_Ret.I_NO.text & " and " & stringyear & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
    Else
       RS.Open "select * from invoiceend where " & stringyear & " and type='" & searchForm & "' order by printorder", con, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    
    RS.Find "TEXT='" + Trim(vs.TextMatrix(I, 0)) + "'"
    If Not RS.EOF Then
        If Trim(RS!DebitorCredit) = Trim("Debit") Then
            If RS!rate > 0 Then
                vs.Col = 1
                frmBookIssueSp_Ret.otherdiscount = frmBookIssueSp_Ret.otherdiscount + ((frmBookIssueSp_Ret.totalamount - frmBookIssueSp_Ret.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                frmBookIssueSp_Ret.otherdiscount = frmBookIssueSp_Ret.otherdiscount - Val(vs.TextMatrix(I, 2))
            End If
        Else
            If RS!rate > 0 Then
                vs.Col = 1
                frmBookIssueSp_Ret.otheramount = frmBookIssueSp_Ret.otheramount + ((frmBookIssueSp_Ret.totalamount - frmBookIssueSp_Ret.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                frmBookIssueSp_Ret.otheramount = frmBookIssueSp_Ret.otheramount - Val(vs.TextMatrix(I, 2))
            End If
        End If
        End If
    
    'RS.close
    Next
    
frmBookIssueSp_Ret.mna.Caption = Format(myround(frmBookIssueSp_Ret.totalamount + frmBookIssueSp_Ret.otheramount - frmBookIssueSp_Ret.totaldiscount - frmBookIssueSp_Ret.otherdiscount, 2), "0.00")
frmBookIssueSp_Ret.mna.Caption = Format(myround(frmBookIssueSp_Ret.totalamount + frmBookIssueSp_Ret.otheramount - frmBookIssueSp_Ret.totaldiscount - frmBookIssueSp_Ret.otherdiscount, 2), "0.00")
Me.Label3 = Format(frmBookIssueSp_Ret.totalamount + frmBookIssueSp_Ret.otheramount - frmBookIssueSp_Ret.totaldiscount - frmBookIssueSp_Ret.otherdiscount, "0.00")
Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
  
  
  
 
End Sub
Sub calc_creditItem()


Dim I As Integer
Critnote.otherdiscount = 0
Critnote.otheramount = 0
    
    
For I = 1 To vs.rows - 1
        
    If RS.State = 1 Then
        RS.close
    End If
    
    If Critnote.edit Then
       RS.Open "select * from CREDITCTMP where invoiceno=" & Critnote.I_NO.text & " and " & stringyear & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
    Else
       RS.Open "select * from invoiceend where " & stringyear & " and type='" & searchForm & "' order by printorder", con, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    
    RS.Find "TEXT='" + Trim(vs.TextMatrix(I, 0)) + "'"
    If Not RS.EOF Then
        If Trim(RS!DebitorCredit) = Trim("Debit") Then
            If RS!rate > 0 Then
                vs.Col = 1
                Critnote.otherdiscount = Critnote.otherdiscount + ((Critnote.totalamount - Critnote.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                Critnote.otherdiscount = Critnote.otherdiscount + Val(vs.TextMatrix(I, 2))
            End If
        Else
            If RS!rate > 0 Then
                vs.Col = 1
                Critnote.otheramount = Critnote.otheramount - ((Critnote.totalamount - Critnote.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                Critnote.otheramount = Critnote.otheramount - Val(vs.TextMatrix(I, 2))
            End If
        End If
        End If
    
    'RS.close
    Next
    
Critnote.mna.Caption = Format(myround(Critnote.totalamount + Critnote.otheramount - Critnote.totaldiscount + Critnote.otherdiscount, 2), "0.00")
Critnote.mna.Caption = Format(myround(Critnote.totalamount + Critnote.otheramount - Critnote.totaldiscount + Critnote.otherdiscount, 2), "0.00")
Me.Label3 = Format(Critnote.totalamount + Critnote.otheramount - Critnote.totaldiscount + Critnote.otherdiscount, "0.00")
Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
'invoice.labelbybank = Me.T3TEXT

End Sub
Sub calc_cash()

Dim I As Integer
countersale.otherdiscount = 0
countersale.otheramount = 0
    
    
For I = 1 To vs.rows - 1
        
    If RS.State = 1 Then
        RS.close
    End If
    
    If countersale.edit Then
       RS.Open "select * from cashctmp where invoiceno=" & countersale.I_NO.text & " and " & stringyear & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
    Else
       RS.Open "select * from invoiceend where " & stringyear & " and type='" & searchForm & "' order by printorder", con, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    
    RS.Find "TEXT='" + Trim(vs.TextMatrix(I, 0)) + "'"
    If Not RS.EOF Then
        If Trim(RS!DebitorCredit) = Trim("Debit") Then
            If RS!rate > 0 Then
                vs.Col = 1
                countersale.otherdiscount = countersale.otherdiscount + ((countersale.totalamount - countersale.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                countersale.otherdiscount = countersale.otherdiscount - Val(vs.TextMatrix(I, 2))
            End If
        Else
            If RS!rate > 0 Then
                vs.Col = 1
                countersale.otheramount = countersale.otheramount + ((countersale.totalamount + countersale.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                countersale.otheramount = countersale.otheramount + Val(vs.TextMatrix(I, 2))
            End If
        End If
        End If
    
    'RS.close
    Next
    

  
  

countersale.mna.Caption = Format(Round(countersale.totalamount + countersale.otheramount - countersale.totaldiscount + countersale.otherdiscount, 2), "0.00")
Me.Label3 = Format(countersale.totalamount + countersale.otheramount - countersale.totaldiscount + countersale.otherdiscount, "0.00")
Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
  
 
End Sub
Sub calc_cashbasil()


T1TEXT.Visible = False
T1.Visible = False
T2TEXT.Visible = False
T2.Visible = False
T3.Visible = False
T3TEXT.Visible = False
Label4.Visible = False
labelbalance.Visible = False



Dim I As Integer
frmBasilSales.otherdiscount = 0
frmBasilSales.otheramount = 0
    
    
For I = 1 To vs.rows - 1
        
    If RS.State = 1 Then
        RS.close
    End If
    
    If frmBasilSales.edit Then
       RS.Open "select * from CASHCTMP_basil where invoiceno=" & frmBasilSales.I_NO.text & " and " & stringyear & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
    Else
       RS.Open "select * from invoiceend where " & stringyear & " and type='" & searchForm & "' order by printorder", con, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    
    RS.Find "TEXT='" + Trim(vs.TextMatrix(I, 0)) + "'"
    If Not RS.EOF Then
        If Trim(RS!DebitorCredit) = Trim("Debit") Then
            If RS!rate > 0 Then
                vs.Col = 1
                frmBasilSales.otherdiscount = frmBasilSales.otherdiscount + ((frmBasilSales.totalamount - frmBasilSales.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                frmBasilSales.otherdiscount = frmBasilSales.otherdiscount - Val(vs.TextMatrix(I, 2))
            End If
        Else
            If RS!rate > 0 Then
                vs.Col = 1
                frmBasilSales.otheramount = frmBasilSales.otheramount + ((frmBasilSales.totalamount + frmBasilSales.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                frmBasilSales.otheramount = frmBasilSales.otheramount + Val(vs.TextMatrix(I, 2))
            End If
        End If
        End If
    
    'RS.close
    Next
    

  
  

frmBasilSales.mna.Caption = Format(Round(frmBasilSales.totalamount + frmBasilSales.otheramount - frmBasilSales.totaldiscount + frmBasilSales.otherdiscount, 2), "0.00")
Me.Label3 = Format(frmBasilSales.totalamount + frmBasilSales.otheramount - frmBasilSales.totaldiscount + frmBasilSales.otherdiscount, "0.00")
Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
  
 
End Sub
Sub calc_cashbasilRet()


T1TEXT.Visible = False
T1.Visible = False
T2TEXT.Visible = False
T2.Visible = False
T3.Visible = False
T3TEXT.Visible = False
Label4.Visible = False
labelbalance.Visible = False




Dim I As Integer
frmBasilSales_Ret.otherdiscount = 0
frmBasilSales_Ret.otheramount = 0
     
    
For I = 1 To vs.rows - 1
        
    If RS.State = 1 Then
        RS.close
    End If
    
    
    If frmBasilSales_Ret.edit Then
       RS.Open "select * from CASHCTMP_basilRet where invoiceno=" & frmBasilSales_Ret.I_NO.text & " and " & stringyear & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
    Else
       RS.Open "select * from invoiceend where " & stringyear & " and type='" & searchForm & "' order by printorder", con, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    
    RS.Find "TEXT='" + Trim(vs.TextMatrix(I, 0)) + "'"
    If Not RS.EOF Then
        If Trim(RS!DebitorCredit) = Trim("Debit") Then
            If RS!rate > 0 Then
                vs.Col = 1
                frmBasilSales_Ret.otherdiscount = frmBasilSales_Ret.otherdiscount + ((frmBasilSales_Ret.totalamount - frmBasilSales_Ret.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                frmBasilSales_Ret.otherdiscount = frmBasilSales_Ret.otherdiscount - Val(vs.TextMatrix(I, 2))
            End If
        Else
            If RS!rate > 0 Then
                vs.Col = 1
                frmBasilSales_Ret.otheramount = frmBasilSales_Ret.otheramount + ((frmBasilSales_Ret.totalamount + frmBasilSales_Ret.totaldiscount) * (Val(vs.TextMatrix(I, 2)) / 100))
            Else
                vs.Col = 2
                frmBasilSales_Ret.otheramount = frmBasilSales_Ret.otheramount + Val(vs.TextMatrix(I, 2))
            End If
        End If
        End If
    
    'RS.close
    Next
    


frmBasilSales_Ret.mna.Caption = Format(Round(frmBasilSales_Ret.totalamount + frmBasilSales_Ret.otheramount - frmBasilSales_Ret.totaldiscount + frmBasilSales_Ret.otherdiscount, 2), "0.00")
Me.Label3 = Format(frmBasilSales_Ret.totalamount + frmBasilSales_Ret.otheramount - frmBasilSales_Ret.totaldiscount + frmBasilSales_Ret.otherdiscount, "0.00")
Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
  
'frmBasilSales_re.mna.Caption = Format(Round(frmBasilSales.totalamount + frmBasilSales.otheramount - frmBasilSales.totaldiscount + frmBasilSales.otherdiscount, 2), "0.00")
'Me.Label3 = Format(frmBasilSales.totalamount + frmBasilSales.otheramount - frmBasilSales.totaldiscount + frmBasilSales.otherdiscount, "0.00")
'Me.labelbalance = Format(Val(Label3) + Val(Me.T1) + Val(Me.T2) - Val(Me.T3TEXT), "0.00")
  
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
   Call CommandReturn_Click
End If

End Sub
Sub iniForm()
Me.top = 1500
Me.Left = 2500

'GROSS.Caption = INVOICE.gr------------------------------------------------


If searchForm = "invoice" Then


            If invoice.edit Then
               Set RS = New ADODB.Recordset
               RS.Open "select * from invoicectmp where " & stringyear & " and invoiceno=" & Val(invoice.I_NO.text) & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  vs.TextMatrix(I, 2) = IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
                
                Else
                        
                        con.Execute ("delete  from invoicectmp WHERE " & stringyear & " and INVOICENO = " + Trim(invoice.I_NO.text))
                        DoEvents
                        con.Execute ("insert into invoicectmp([INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],[saleType]) " & _
                        "  select [INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],[saleType]" & _
                        " from invoicec where  " & stringyear & " and INVOICENO = " + Trim(invoice.I_NO.text))
                        DoEvents

                End If
            
            
            Else
              
              
              Set RS = New ADODB.Recordset
              RS.Open "select * from INVOICEEND where " & stringyear & " and type='" & searchForm & "' ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
              If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  'vs.TextMatrix(I, 2) = 0 'IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            End If
            
            
            End If
            
            
            
            '===========================================================================
            '---------------------------------------------------------------------------
            If RS.State = 1 Then RS.close
            RS.Open "select * from invoicea where " & stringyear & " and " & _
            " invoiceno=" + Trim(invoice.I_NO.text), con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not RS.EOF Then
               Me.T1 = Format(myround(RS!txt1a, 2), "0.00")
               Me.T1TEXT = IIf(IsNull(RS!txt1), "", RS!txt1)
               Me.T2 = Format(myround(RS!txt2a, 2), "0.00")
               Me.T2TEXT = IIf(IsNull(RS!txt2), "", RS!txt2)
               Me.T3TEXT = Format(myround(RS!baa, 2), "0.00")
            End If
            RS.close
            
            
            
            calc
            setWidth
ElseIf searchForm = "invoiceblue" Then


            If frmInvoice_blueprint.edit Then
               Set RS = New ADODB.Recordset
               RS.Open "select * from invoicectmp_blue where " & stringyear & " and invoiceno=" & Val(frmInvoice_blueprint.I_NO.text) & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  vs.TextMatrix(I, 2) = IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
                
            Else
                    
                con.Execute ("delete  from invoicectmp_blue WHERE " & stringyear & " and INVOICENO = " + Trim(frmInvoice_blueprint.I_NO.text))
                DoEvents
                con.Execute ("insert into invoicectmp_blue([INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],[saleType]) " & _
                "  select [INVOICENO] , [INVOICEDATE], [Genledger], [SUBLEDGER], [GAMOUNT], [rate], [amount], [DEBITORCREDIT], [Text], [RYN], [Fyear], [setupid],[saleType]" & _
                " from invoicec_blue where  " & stringyear & " and INVOICENO = " + Trim(frmInvoice_blueprint.I_NO.text))
                DoEvents

            End If
            
            
            Else


              Set RS = New ADODB.Recordset
              RS.Open "select * from INVOICEEND where " & stringyear & " and type='" & searchForm & "' ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
              If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  'vs.TextMatrix(I, 2) = 0 'IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            End If
            
            End If
            
            
            
            '===========================================================================
            '---------------------------------------------------------------------------
            If RS.State = 1 Then RS.close
            RS.Open "select * from invoicea_blue where " & stringyear & " and " & _
            " invoiceno=" + Trim(frmInvoice_blueprint.I_NO.text), con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not RS.EOF Then
               Me.T1 = Format(myround(RS!txt1a, 2), "0.00")
               Me.T1TEXT = IIf(IsNull(RS!txt1), "", RS!txt1)
               Me.T2 = Format(myround(RS!txt2a, 2), "0.00")
               Me.T2TEXT = IIf(IsNull(RS!txt2), "", RS!txt2)
               Me.T3TEXT = Format(myround(RS!baa, 2), "0.00")
            End If
            RS.close
            
            
            
            calc_invblue
            setWidth


ElseIf searchForm = "invoice_sp" Then


            If frmBookIssueSp.edit Then
               Set RS = New ADODB.Recordset
               RS.Open "select * from invoicectmp_sp where " & stringyear & " and username ='" & UserName & "' and invoiceno=" & Val(frmBookIssueSp.I_NO.text) & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  vs.TextMatrix(I, 2) = IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            Else
             
              '''''
              '''''
              '''''
                         
            End If
            
            
            Else


              Set RS = New ADODB.Recordset
              RS.Open "select * from INVOICEEND where " & stringyear & " and type='" & searchForm & "' ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
              If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  'vs.TextMatrix(I, 2) = 0 'IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            End If
            
            
            End If
            
            
            
            '===========================================================================
            '---------------------------------------------------------------------------
            If RS.State = 1 Then RS.close
            RS.Open "select * from invoicea_sp where " & stringyear & " and " & _
            " invoiceno=" + Trim(frmBookIssueSp.I_NO.text), con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not RS.EOF Then
               Me.T1 = Format(myround(RS!txt1a, 2), "0.00")
               Me.T1TEXT = IIf(IsNull(RS!txt1), "", RS!txt1)
               Me.T2 = Format(myround(RS!txt2a, 2), "0.00")
               Me.T2TEXT = IIf(IsNull(RS!txt2), "", RS!txt2)
               Me.T3TEXT = Format(myround(RS!baa, 2), "0.00")
            End If
            RS.close
            
            
            
            calc_invoice_sp
            setWidth

ElseIf searchForm = "invoice_spret" Then


            If frmBookIssueSp_Ret.edit Then
               Set RS = New ADODB.Recordset
               RS.Open "select * from invoicectmp_spret where " & stringyear & " and username='" & UserName & "' and invoiceno=" & Val(frmBookIssueSp_Ret.I_NO.text) & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  vs.TextMatrix(I, 2) = IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            Else
                  ''''GoTo invoicec_spret:
                  ''''
                  ''''
            End If
            
            
            Else
             
            
              Set RS = New ADODB.Recordset
              RS.Open "select * from INVOICEEND where " & stringyear & " and type='" & searchForm & "' ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
              If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  'vs.TextMatrix(I, 2) = 0 'IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            End If
            
            
            End If
            
            
            
            '===========================================================================
            '---------------------------------------------------------------------------
            If RS.State = 1 Then RS.close
            RS.Open "select * from invoicea_spret where " & stringyear & " and " & _
            " invoiceno=" + Trim(frmBookIssueSp_Ret.I_NO.text), con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not RS.EOF Then
               Me.T1 = Format(myround(RS!txt1a, 2), "0.00")
               Me.T1TEXT = IIf(IsNull(RS!txt1), "", RS!txt1)
               Me.T2 = Format(myround(RS!txt2a, 2), "0.00")
               Me.T2TEXT = IIf(IsNull(RS!txt2), "", RS!txt2)
               Me.T3TEXT = Format(myround(RS!baa, 2), "0.00")
            End If
            RS.close
            
            
            
            calc_invoice_spRet
            setWidth


ElseIf searchForm = "credititem" Then


            If Critnote.edit Then
               
               Set RS = New ADODB.Recordset
               RS.Open "select * from CREDITCTMP where " & stringyear & " and invoiceno=" & Val(Critnote.I_NO.text) & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
               If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  vs.TextMatrix(I, 2) = IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
                
             Else
                   ' CREDITCtmp creation start
                    DoEvents
                    con.Execute ("delete from CREDITCtmp where INVOICENO = " & Critnote.I_NO & " and " & stringyear)
                    con.Execute ("insert into CREDITCtmp(INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate," & _
                    "AMOUNT,DEBITORCREDIT,TEXT,RYN,fyear," & _
                    "setupid)  select INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate," & _
                    "AMOUNT,DEBITORCREDIT,TEXT,RYN,fyear," & _
                    "setupid from CREDITC where " & stringyear & " and INVOICENO = " + Trim(Critnote.I_NO.text))
                    DoEvents
                    ' CREDITTMP creation end

             End If
            
            
            Else
              

              
              Set RS = New ADODB.Recordset
              RS.Open "select * from invoiceEnd where " & stringyear & " and type='" & searchForm & "' ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
              If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  'vs.TextMatrix(I, 2) = 0 'IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            End If
            
            
            End If
            
            
            
            '===========================================================================
            '---------------------------------------------------------------------------
            If RS.State = 1 Then RS.close
            RS.Open "select * from CREDITA where " & stringyear & " and " & _
            " invoiceno=" + Trim(Critnote.I_NO.text), con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not RS.EOF Then
               Me.T1 = Format(myround(RS!txt1a & "", 2), "0.00")
               Me.T1TEXT = IIf(IsNull(RS!txt1), "", RS!txt1)
               Me.T2 = Format(myround(RS!txt2a & "", 2), "0.00")
               Me.T2TEXT = IIf(IsNull(RS!txt2), "", RS!txt2)
               Me.T3TEXT = Format(myround(RS!baa & "", 2), "0.00")
            End If
            RS.close
            
            
            
            calc_creditItem
            setWidth



ElseIf searchForm = "cash" Then
                  

            If countersale.edit Then
               
            
               Set RS = New ADODB.Recordset
               RS.Open "select * from cashctmp where " & stringyear & " and invoiceno=" & Val(countersale.I_NO.text) & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  vs.TextMatrix(I, 2) = IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            Else
                    DoEvents
                    con.Execute "Delete  from CASHCTMP where  " & stringyear & " and INVOICENO = " & Val(countersale.I_NO.text) & ""
                    DoEvents
                    con.Execute ("insert into CASHCTMP(INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate,AMOUNT,DEBITORCREDIT,[TEXT],RYN,Fyear,setupid)  select INVOICENO,INVOICEDATE,GENLEDGER,SUBLEDGER,GAMOUNT,Rate,AMOUNT,DEBITORCREDIT,[TEXT],RYN,Fyear,setupid from CASHC where " & stringyear & " and INVOICENO = " + Trim(countersale.I_NO.text))

            End If
            
            
            Else
              Set RS = New ADODB.Recordset
              RS.Open "select * from invoiceend where type='" & searchForm & "' and " & stringyear & " ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
              If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  'vs.TextMatrix(I, 2) = ""
                  RS.MoveNext
                Next
            
            End If
            
            
            End If
            
            
            
            '===========================================================================
            '---------------------------------------------------------------------------
            If RS.State = 1 Then RS.close
            RS.Open "select * from cashA where " & stringyear & " and " & _
            " invoiceno=" + Trim(countersale.I_NO.text), con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not RS.EOF Then
               Me.T1 = Format(myround(RS!txt1a & "", 2), "0.00")
               Me.T1TEXT = IIf(IsNull(RS!txt1), "", RS!txt1)
               Me.T2 = Format(myround(RS!txt2a & "", 2), "0.00")
               Me.T2TEXT = IIf(IsNull(RS!txt2), "", RS!txt2)
               
               Me.T3TEXT = Format(myround(RS!baa & "", 2), "0.00")
            End If
            RS.close
            
            calc_cash
            setWidth


ElseIf searchForm = "cashbasil" Then
                  
            
            If frmBasilSales.edit Then
               Set RS = New ADODB.Recordset
               RS.Open "select * from CASHCTMP_basil where " & stringyear & " and invoiceno=" & Val(frmBasilSales.I_NO.text) & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  vs.TextMatrix(I, 2) = IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            End If
            
            
            Else
              Set RS = New ADODB.Recordset
              RS.Open "select * from invoiceend where type='" & searchForm & "' and " & stringyear & " ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
              If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  '''vs.TextMatrix(I, 2) = "" 'IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            End If
            
            
            End If
            
            
            
            '===========================================================================
            '---------------------------------------------------------------------------
            If RS.State = 1 Then RS.close
            RS.Open "select * from CASHA_basil where " & stringyear & " and " & _
            " invoiceno=" + Trim(frmBasilSales.I_NO.text), con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not RS.EOF Then
               Me.T1 = Format(myround(RS!txt1a & "", 2), "0.00")
               Me.T1TEXT = IIf(IsNull(RS!txt1), "", RS!txt1)
               Me.T2 = Format(myround(RS!txt2a & "", 2), "0.00")
               Me.T2TEXT = IIf(IsNull(RS!txt2), "", RS!txt2)
               Me.T3TEXT = Format(myround(RS!baa & "", 2), "0.00")
            End If
            RS.close
            
            calc_cashbasil
            setWidth

ElseIf searchForm = "cashbasilret" Then
                  
            
            If frmBasilSales_Ret.edit Then
               Set RS = New ADODB.Recordset
               RS.Open "select * from CASHCTMP_basilret where " & stringyear & " and invoiceno=" & Val(frmBasilSales_Ret.I_NO.text) & " order by autoid", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  vs.TextMatrix(I, 2) = IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            End If
            
            
            Else
              Set RS = New ADODB.Recordset
              RS.Open "select * from invoiceend where type='" & searchForm & "' and " & stringyear & " ORDER BY PRINTORDER", con, adOpenKeyset, adLockReadOnly, adCmdText
              If RS.EOF = False Then
            
                vs.rows = RS.RecordCount + 1
                For I = 1 To RS.RecordCount
                
                  vs.TextMatrix(I, 0) = RS!text
                  vs.TextMatrix(I, 1) = IIf(RS!rate = 0, "", RS!rate)
                  '''vs.TextMatrix(I, 2) = "" 'IIf(RS!amount = 0, "", RS!amount)
                  RS.MoveNext
                Next
            
            End If
            
            
            End If
            
            
            '===========================================================================
            '---------------------------------------------------------------------------
            If RS.State = 1 Then RS.close
            RS.Open "select * from CASHA_basilRet where " & stringyear & " and " & _
            " invoiceno=" + Trim(frmBasilSales_Ret.I_NO.text), con, adOpenKeyset, adLockReadOnly, adCmdText
            If Not RS.EOF Then
               Me.T1 = Format(myround(RS!txt1a & "", 2), "0.00")
               Me.T1TEXT = IIf(IsNull(RS!txt1), "", RS!txt1)
               Me.T2 = Format(myround(RS!txt2a & "", 2), "0.00")
               Me.T2TEXT = IIf(IsNull(RS!txt2), "", RS!txt2)
               Me.T3TEXT = Format(myround(RS!baa & "", 2), "0.00")
            End If
            RS.close
            
            calc_cashbasilRet
            setWidth

End If


End Sub
Sub setWidth()
 
 vs.Cols = 3
 vs.FormatString = "Ledger|>Rate(%)if any|>Amount(rs)"
 vs.ColWidth(0) = 4000
 vs.ColWidth(1) = 1200
 vs.ColWidth(2) = 1200

End Sub
Private Sub Form_Load()
  iniForm
  BackColorFrom Me
  k10 = 0
  
  If (LCase(UserName) = "admin" Or LCase(UserName) = "v") Then
     cmdref.Enabled = True
     Label1(20).Visible = True
     Label1(21).Visible = True
  Else
     cmdref.Enabled = False
  End If
  

End Sub

Private Sub Form_Unload(cancel As Integer)
k10 = 0
End Sub

Private Sub T1_GotFocus()
VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.text)
End Sub
Private Sub T1_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     sendkeys "{tab}"
  If searchForm = "credititem" Then
      calc_creditItem
  ElseIf searchForm = "cash" Then
       calc_cash
  ElseIf searchForm = "cashbasil" Then
       calc_cashbasil
  ElseIf searchForm = "cashbasilret" Then
       calc_cashbasilRet
  ElseIf searchForm = "invoice" Then
       calc
  ElseIf searchForm = "invoice_sp" Then
       calc_invoice_sp
  ElseIf searchForm = "invoice_spret" Then
       calc_invoice_spRet
  End If
  End If
End Sub
Private Sub T1TEXT_GotFocus()
On Error Resume Next
VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.text)
End Sub
Private Sub T1TEXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then T1.SetFocus
End Sub

Private Sub T2_GotFocus()
VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.text)
End Sub
Private Sub Timer1_Timer()
Static L As Integer
Static kk1 As Integer



If (frt = "No") Then

 If din_ = 1 Then
 
    Label1(20).Caption = "Freight not allowed..."
    Label1(21).Caption = "Freight not allowed..."
    
    
    If L = 0 Then
        Label1(20).ForeColor = vbWhite
        Label1(21).ForeColor = vbRed
        L = 1
        din_ = din_ + 1
        Exit Sub
    ElseIf L = 1 Then
        Label1(20).ForeColor = vbBlue
        Label1(21).ForeColor = vbRed
        L = 0
        din_ = din_ + 1
        Exit Sub
    End If
  End If
  
    
    
Else
    If din_ <= 100 Then din_ = 2
    Label1(20).Caption = ""
    Label1(21).Caption = ""
    
End If





If (postage = "No") Then

    If din_ = 2 Then
    
    Label1(20).Caption = "Postage not allowed..."
    Label1(21).Caption = "Postage not allowed..."
    
    
    If L = 0 Then
        Label1(20).ForeColor = vbWhite
        Label1(21).ForeColor = vbRed
        L = 1
        din_ = din_ + 1
        
        If din_ <= 100 Then din_ = 1
        
        Exit Sub
    ElseIf L = 1 Then
        Label1(20).ForeColor = vbBlue
        Label1(21).ForeColor = vbRed
        L = 0
        din_ = din_ + 1
        If din_ <= 100 Then din_ = 1
        Exit Sub
    End If
    
    End If

Else
    Label1(20).Caption = ""
    Label1(21).Caption = ""
    If din_ <= 100 Then din_ = 1




End If
    

    



End Sub

Private Sub T2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  
  
 
  If searchForm = "credititem" Then
     Me.CommandReturn.SetFocus
  Else
     sendkeys "{tab}"
  End If
  
  
  If searchForm = "credititem" Then
    calc_creditItem
  ElseIf searchForm = "cash" Then
     calc_cash
  ElseIf searchForm = "cashbasil" Then
     calc_cashbasil
  ElseIf searchForm = "cashbasilret" Then
     calc_cashbasilRet
  ElseIf searchForm = "invoice" Then
   calc
  ElseIf searchForm = "invoice_sp" Then
   calc_invoice_sp
  ElseIf searchForm = "invoice_spret" Then
   calc_invoice_spRet
  End If
  
  End If

End Sub

Private Sub T2TEXT_GotFocus()
VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.text)
End Sub

Private Sub T2TEXT_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then T2.SetFocus
End Sub

Private Sub T3TEXT_GotFocus()
VB.Screen.ActiveForm.ActiveControl.SelLength = Len(VB.Screen.ActiveControl.text)
End Sub

Private Sub T3TEXT_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    
    T3TEXT.text = Format(T3TEXT.text, "0.00")
    Me.CommandReturn.SetFocus
    
  If searchForm = "credititem" Then
    calc_creditItem
  ElseIf searchForm = "cash" Then
     calc_cash
  ElseIf searchForm = "cashbasil" Then
     calc_cashbasil
  ElseIf searchForm = "cashbasilret" Then
     calc_cashbasilRet
  
  ElseIf searchForm = "invoice" Then
   calc
  ElseIf searchForm = "invoice_sp" Then
   calc_invoice_sp
  ElseIf searchForm = "invoice_spret" Then
   calc_invoice_spRet
  
  End If
    
Else
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
End If



End Sub
Private Sub T3TEXT_LostFocus()

If searchForm = "cash" Then
   countersale.labelbybank = T3TEXT
End If
 
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
   vs.TextMatrix(vs.RowSel, 2) = ""
End If
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

If k10 = 0 Then
    k10 = k10 + 1
    Exit Sub
End If

If KeyCode = 13 Then
   
   
''  If vs.Col = 2 Then
''
''  If searchForm = "credititem" Then
''     calc_creditItem
''  ElseIf searchForm = "cash" Then
''     calc_cash
''  ElseIf searchForm = "cashbasil" Then
''     calc_cashbasil
''  ElseIf searchForm = "cashbasilret" Then
''     calc_cashbasilRet
''  ElseIf searchForm = "invoice" Then
''     calc
''  ElseIf searchForm = "invoice_sp" Then
''     calc_invoice_sp
''  ElseIf searchForm = "invoice_spret" Then
''     calc_invoice_spRet
''  End If
sendkeys "{down}"
End If
'End If
End Sub

Private Sub vs_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
   
   sendkeys "{down}"
   If (vs.rows - 1) = vs.Row Then

   'If T1TEXT.Visible = True Then
      T1TEXT.SetFocus
   '    MsgBox "1"
   'Else
   '   Commandreturn.SetFocus
   '   MsgBox "2"
   'End If

   End If
   
End If
End Sub
Private Sub vs_LostFocus()
If vs.Col = 2 Then
     
  If searchForm = "credititem" Then
     calc_creditItem
  ElseIf searchForm = "cash" Then
     calc_cash
  ElseIf searchForm = "cashbasil" Then
     calc_cashbasil
  ElseIf searchForm = "cashbasilret" Then
     calc_cashbasilRet
  ElseIf searchForm = "invoice" Then
     calc
  ElseIf searchForm = "invoice_sp" Then
     calc_invoice_sp
  ElseIf searchForm = "invoice_spret" Then
     calc_invoice_spRet
  ElseIf searchForm = "invoiceblue" Then
     calc_invblue
  
  End If
  
End If
End Sub
Private Sub vs_SelChange()
    
If searchForm = "invoice" Then
    
    If vs.Col = 2 Then
       vs.Editable = flexEDKbd
    
    If Right(session, 2) >= 20 Then
    If vs.TextMatrix(vs.RowSel, 0) = "LESS FREIGHT" Then
        If invoice.lblPartyfrt.Caption = "No" Then
           vs.Editable = flexEDNone
        End If
    End If
    
    If vs.TextMatrix(vs.RowSel, 0) = "ADD POSTAGE" Then
        If invoice.lblPostage.Caption = "No" Then
           vs.Editable = flexEDNone
        End If
    End If
    
    End If
    
    End If
    
End If
    
End Sub

