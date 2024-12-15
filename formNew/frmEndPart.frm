VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmEndPart 
   Caption         =   "End Part"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   11040
   Begin VB.Frame panel 
      Caption         =   "End Part"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7170
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10980
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   240
         TabIndex        =   8
         Top             =   6165
         Width           =   7935
         Begin VB.CommandButton cmdSearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   5085
            Picture         =   "frmEndPart.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   45
            Width           =   1365
         End
         Begin VB.CommandButton cmdExit_12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   6465
            Picture         =   "frmEndPart.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   30
            Width           =   1410
         End
         Begin VB.CommandButton cmdEdit_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   3795
            Picture         =   "frmEndPart.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   45
            Width           =   1275
         End
         Begin VB.CommandButton cmdDelete_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   2520
            Picture         =   "frmEndPart.frx":1BD5
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   45
            Width           =   1275
         End
         Begin VB.CommandButton cmdSave_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   1290
            Picture         =   "frmEndPart.frx":27B9
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   45
            Width           =   1230
         End
         Begin VB.CommandButton cmdAdd_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   45
            Picture         =   "frmEndPart.frx":339D
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   45
            Width           =   1230
         End
      End
      Begin VB.TextBox txtRate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   7
         Top             =   2205
         Width           =   1335
      End
      Begin VB.ComboBox cboDRCR 
         Height          =   315
         ItemData        =   "frmEndPart.frx":3F81
         Left            =   3060
         List            =   "frmEndPart.frx":3F8B
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2535
         Width           =   1365
      End
      Begin VB.ComboBox cboSubledger 
         Height          =   315
         Left            =   3060
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1500
         Width           =   3930
      End
      Begin VB.ComboBox cboGenDesc 
         Height          =   315
         Left            =   3060
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1140
         Width           =   3930
      End
      Begin VB.ComboBox cboGenContra 
         Height          =   315
         Left            =   3060
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   765
         Width           =   3930
      End
      Begin VB.TextBox txtDesc20 
         Height          =   285
         Left            =   3060
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1860
         Width           =   3900
      End
      Begin VB.TextBox TextInvePrintOrder 
         Height          =   315
         Left            =   3060
         MaxLength       =   6
         TabIndex        =   1
         Top             =   405
         Width           =   1335
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   2985
         Left            =   225
         TabIndex        =   15
         Top             =   2970
         Width           =   10710
         _cx             =   18891
         _cy             =   5265
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
         BackColorBkg    =   -2147483636
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
         Editable        =   0
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   1005
         Left            =   225
         Top             =   6120
         Width           =   7980
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Print Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   495
         Width           =   1380
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate % (if any)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Debit/Credit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   2475
         Width           =   3015
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Gen.ledger Desc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1155
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub.ledger Desc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   1500
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Contra Gen. Ledger Desc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   17
         Top             =   795
         Width           =   2955
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "20 char.Text-->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1815
         Width           =   2985
      End
      Begin VB.Image imgFile 
         Height          =   240
         Left            =   5820
         Picture         =   "frmEndPart.frx":3F9E
         Top             =   2250
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6810
         Picture         =   "frmEndPart.frx":40E8
         Top             =   2070
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmEndPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f_grid As New ADODB.Recordset

Private Sub cboGenDesc_Click()

cboSubledger.Enabled = True

Set RS = New ADODB.Recordset
RS.Open "select * from sledger where gledger='" + Trim(cboGenDesc.Text) + "' and " & stringyear & "", con, adOpenDynamic, adLockReadOnly
cboSubledger.Clear

If Not RS.BOF Then
    
    RS.MoveFirst
    Do While Not RS.EOF
        Me.cboSubledger.AddItem RS(1)
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
Else
   cboSubledger.Enabled = False
End If

End Sub
Sub fillGrid()

If f_grid.State = 1 Then f_grid.close

f_grid.Open "select CGENLEDGER,GENLEDGER,SUBLEDGER,TEXT,Rate,DEBITORCREDIT,PrintOrder from INVOICEEND where type='" & popupvalue5 & "' and " & stringyear & "  order by PrintOrder", con, adOpenForwardOnly, adLockReadOnly
Set vs.DataSource = f_grid

vs.ColWidth(1) = 1500

For I = 1 To vs.Rows - 1
vs.Cell(flexcpPicture, I, 1) = imgFile
Next


End Sub
Private Sub cmdAdd_1_Click()

clearFrom Me
TextInvePrintOrder.SetFocus

fillGrid

End Sub

Private Sub cmdDelete_3_Click()

If MsgBox("Want to Delete ?", vbYesNo + vbQuestion) = vbYes Then
  con.Execute "delete from [INVOICEEND] where [PrintOrder]=" & TextInvePrintOrder & " and " & stringyear & " and type='" & popupvalue5 & "'"
  Call cmdAdd_1_Click
End If

End Sub

Private Sub cmdEdit_4_Click()

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdSave_2_Click()



If TextInvePrintOrder = "" Then
  MsgBox "Enter Print Order No..", vbCritical
  TextInvePrintOrder.SetFocus
  Exit Sub
End If

If Trim(cboGenContra) = "" Then
   MsgBox "Contra Gen. Ledger Desc.", vbCritical
   cboGenContra.SetFocus
   Exit Sub
End If


If Trim(cboGenDesc) = "" Then
 MsgBox "Gen.ledger Desc.", vbCritical
 cboGenDesc.SetFocus
 Exit Sub
End If

If Trim(txtDesc20) = "" Then
   MsgBox "20 char.Text-->", vbCritical
   txtDesc20.SetFocus
   Exit Sub
End If

If cboDRCR.Text = "" Then
   MsgBox "Enter Debit/Credit ..", vbCritical
   cboDRCR.SetFocus
   Exit Sub
End If


If RS.State = 1 Then RS.close
RS.Open "SELECT [CGENLEDGER],[GENLEDGER],[SUBLEDGER],[TEXT],[RATE],[DEBITORCREDIT],[PrintOrder],setupid,fyear,Type from [INVOICEEND] " & _
"where [PrintOrder]=" & TextInvePrintOrder & " and " & stringyear & " and type='" & Trim(popupvalue5) & "'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then RS.AddNew

RS!PRINTORDER = TextInvePrintOrder
RS!CGENLEDGER = Trim(cboGenContra)
RS!Genledger = Trim(cboGenDesc)
RS!SUBLEDGER = Trim(cboSubledger)
RS!Text = Trim(txtDesc20)
RS!rate = Val(txtRate)
RS!DebitorCredit = cboDRCR.Text
RS!setupid = setupid
RS!fyear = session
RS!Type = Trim(LCase(popupvalue5))
RS.update
  
If popupvalue5 = "CREDITITEM" Then
con.Execute "update CREDITC set DEBITORCREDIT='" & cboDRCR.Text & "' where (GENLEDGER='" & Trim(cboGenDesc) & "' and Text='" & Trim(txtDesc20) & "')"
ElseIf popupvalue5 = "INVOICE_SPRET" Then
con.Execute "update INVOICEC_spRet set DEBITORCREDIT='" & cboDRCR.Text & "' where (Text='" & Trim(txtDesc20) & "')"
End If
  
MsgBox "Data Saved ...", vbInformation

Call cmdAdd_1_Click




End Sub

Private Sub Form_Load()
     
  Me.Left = 100
  Me.Top = 100
  Me.Width = 11200
  Me.Height = 7800
   
     
  BackColorFrom Me
  
  panel.Caption = popupvalue5 & " End Part"
  'popupvalue5 = ""
     
  If RS.State = 1 Then RS.close
  RS.Open "select gledger from gledger where " & stringyear & " order by gledger", con, adOpenDynamic, adLockReadOnly
    If Not RS.EOF Then
        Do While Not RS.EOF
            
            Me.cboGenContra.AddItem RS!gledger
            Me.cboGenDesc.AddItem RS!gledger
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    
    fillGrid
    
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub vs_DblClick()
  
  cmdSave_2.Enabled = False
  cmdEdit_4.Enabled = True
  cmdDelete_3.Enabled = True
  
  TextInvePrintOrder.Text = vs.TextMatrix(vs.RowSel, 7)
  cboGenContra.Text = vs.TextMatrix(vs.RowSel, 1)
  cboGenDesc.Text = vs.TextMatrix(vs.RowSel, 2)
  cboSubledger.Text = vs.TextMatrix(vs.RowSel, 3)
  txtDesc20.Text = vs.TextMatrix(vs.RowSel, 4)
  
  txtRate.Text = vs.TextMatrix(vs.RowSel, 5)
  
  cboDRCR.Text = vs.TextMatrix(vs.RowSel, 6)
  
  
  
  
End Sub
