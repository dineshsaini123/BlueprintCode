VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMS 
   Caption         =   "SMS Status.."
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10950
   Icon            =   "frmSMS.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   10950
   Begin VB.OptionButton Option2_sp 
      Caption         =   "Specimen"
      Height          =   330
      Left            =   6210
      TabIndex        =   7
      Top             =   315
      Width           =   1140
   End
   Begin VB.OptionButton Option1_sale 
      Caption         =   "Sale"
      Height          =   330
      Left            =   5085
      TabIndex        =   6
      Top             =   315
      Value           =   -1  'True
      Width           =   1050
   End
   Begin VB.CheckBox Check1_sended 
      Caption         =   "Total Sended SMS"
      Height          =   330
      Left            =   3015
      TabIndex        =   5
      Top             =   315
      Width           =   1905
   End
   Begin VB.CommandButton cmdref 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   540
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   135
      Width           =   1095
   End
   Begin VB.CommandButton cmdSendSMS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Send SMS"
      Height          =   510
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   225
      Width           =   1185
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7320
      Left            =   90
      TabIndex        =   0
      Top             =   1125
      Width           =   10725
      _cx             =   18918
      _cy             =   12912
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   7917545
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   -2147483647
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin MSComCtl2.DTPicker Dates 
      Height          =   315
      Left            =   135
      TabIndex        =   3
      Top             =   405
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   71827457
      CurrentDate     =   39500
   End
   Begin VB.Label lblTotal 
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   8505
      Width           =   825
   End
   Begin VB.Shape Shape1 
      Height          =   600
      Left            =   4995
      Top             =   180
      Width           =   2445
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SMS Date "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   90
      Width           =   1455
   End
End
Attribute VB_Name = "frmSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub cmdRef_Click()
fillData_VS
End Sub
Private Sub cmdSendSMS_Click()
  
  For I = 1 To vs.Rows - 1
  
  If vs.TextMatrix(I, 5) = "n" Then
     
     smsSend vs.TextMatrix(I, 0), vs.TextMatrix(I, 3), vs.TextMatrix(I, 4)
     Sleep (1000)
     
      vs.TextMatrix(I, 5) = "Sent..."
      For k1 = 0 To 5
             vs.Cell(flexcpBackColor, I, k1) = vbGreen
              DoEvents
              DoEvents
       Next

     
  End If
 Next

    

End Sub

Sub fillData_VS()
       Dim rs_ As New ADODB.Recordset
       Dim k1 As Integer
       Dim k2 As Integer
       vs.Clear
       If rs_.State = 1 Then rs_.close
       
       
       If Option1_sale.value = True Then
       
                If Check1_sended.value = 0 Then
                   rs_.Open "select InvoiceNo,InvoiceDate,Subledger,Mobile,RandomId,SMSSend as [SMS Status] from invoicea where smsdate =convert(smalldatetime,'" & Dates.value & "',103) and SMSSend='n' and mobile is not null", con
                Else
                  rs_.Open "select InvoiceNo,InvoiceDate,Subledger,Mobile,RandomId,SMSSend as [SMS Status] from invoicea where smsdate =convert(smalldatetime,'" & Dates.value & "',103) and SMSSend='y' and mobile is not null", con
                End If
       Else
       
                If Check1_sended.value = 0 Then
                   rs_.Open "select InvoiceNo,InvoiceDate,Agentname as Subledger,Mobile,RandomId,SMSSend as [SMS Status] from invoicea_sp where smsdate =convert(smalldatetime,'" & Dates.value & "',103) and SMSSend='n' and mobile is not null", con
                Else
                  rs_.Open "select InvoiceNo,InvoiceDate,Agentname as Subledger,Mobile,RandomId,SMSSend as [SMS Status] from invoicea_sp where smsdate =convert(smalldatetime,'" & Dates.value & "',103) and SMSSend='y' and mobile is not null", con
                End If
       
       End If
       
       
       
       If rs_.RecordCount > 0 Then
          cmdSendSMS.Enabled = True
       Else
          cmdSendSMS.Enabled = False
       End If
       
       vs.Rows = 1
       lblTotal.Caption = 0
       k2 = 0
       For k1 = 1 To rs_.RecordCount
       
      If rs_!mobile <> "" Then
       
       vs.Rows = vs.Rows + 1
       k2 = k2 + 1
       vs.TextMatrix(k2, 0) = rs_!invoiceNo
       vs.TextMatrix(k2, 1) = rs_!INVOICEDATE
       vs.TextMatrix(k2, 2) = rs_!SUBLEDGER
       vs.TextMatrix(k2, 3) = rs_!mobile & ""
       vs.TextMatrix(k2, 4) = rs_!randomId & ""
       vs.TextMatrix(k2, 5) = rs_(5)
      End If
      ' lblTotal.Caption = Val(lblTotal.Caption) + 1
       
       rs_.MoveNext
       
       Next
       
       
vs.FormatString = "InvoiceNo|InvoiceDate|Party|Mobile|RandomId|SMS Status"
       
 lblTotal.Caption = "Total : " & vs.Rows - 1
       
vs.ColWidth(0) = 1100
vs.ColWidth(1) = 1100
vs.ColWidth(2) = 4200
vs.ColWidth(3) = 1200
vs.ColWidth(4) = 1200
vs.ColWidth(5) = 1200


       
       
End Sub


Private Sub Dates_LostFocus()
fillData_VS
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 9180
Me.Width = 11000

Dates.value = Format(Date, "dd/MM/yyyy")

cmdRef_Click

End Sub
