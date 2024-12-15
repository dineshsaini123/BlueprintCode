VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmBookReceipt 
   Caption         =   "Book Receipt"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   18495
   WindowState     =   2  'Maximized
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   2400
      TabIndex        =   9
      Top             =   7200
      Width           =   8835
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
         Height          =   840
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1230
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
         Height          =   840
         Left            =   1275
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1230
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
         Height          =   840
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   1230
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
         Height          =   840
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   1230
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
         Height          =   840
         Left            =   7500
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1230
      End
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
         Height          =   840
         Left            =   6255
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9825
      TabIndex        =   8
      Top             =   6600
      Width           =   1215
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4965
      Left            =   1650
      TabIndex        =   3
      Top             =   1575
      Width           =   9765
      _cx             =   17224
      _cy             =   8758
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VB.TextBox txtRem 
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   1050
      Width           =   6690
   End
   Begin MSMask.MaskEdBox RecDate 
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtRecNo 
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Remarks :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   450
      TabIndex        =   7
      Top             =   1050
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "Rec. Date"
      Height          =   240
      Left            =   3375
      TabIndex        =   6
      Top             =   675
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Receipt No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   450
      TabIndex        =   5
      Top             =   600
      Width           =   1290
   End
End
Attribute VB_Name = "frmBookReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim editData As Boolean
Private Sub cmdAdd_1_Click()

MaxRecNo
RecDate.Text = Format(Date, "dd/MM/yyyy")
RecDate.SetFocus
txtRem = ""
setwidth

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False
txtRem.SetFocus

End Sub

Private Sub cmdDelete_3_Click()
 If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
    CON.Execute "delete from ProductReceipt where " & stringyear & " and  RecNo=" & txtRecNo & " and " & stringyear & ""
    cmdAdd_1_Click
 End If
End Sub

Private Sub cmdEdit_4_Click()

cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = True
cmdSave_2.Enabled = True

editData = True
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub
Sub searchData()
  
  setwidth
  
  i = 1
  K = 1
  
  Dim rs1 As New ADODB.Recordset
  
  
  If rs.State = 1 Then rs.Close
  rs.Open "SELECT RecNo,RecDate,PCode," & _
  "QUANTITY,remarks FROM ProductReceipt " & _
  " where " & stringyear & " and  RecNo=" & txtRecNo & "", CON, adOpenKeyset, adLockReadOnly
  If rs.EOF = False Then
     
     cmdDelete_3.Enabled = False
     cmdSave_2.Enabled = False
     
     RecDate.Text = rs(1)
     txtRem = rs(2)
     While rs.EOF = False
          
          
          
        If rs1.State = 1 Then rs1.Close
       

           vs.TextMatrix(i, 0) = K
           vs.TextMatrix(i, 1) = rs(2)
         
            rs1.Open "select TypeofProduct ,rulling," & _
            "NoOfPages ,ProductQuality from copymaster where " & stringyear & " and  bookno='" & rs(2) & "'", CON, adOpenKeyset, adLockReadOnly
            If rs1.EOF = False Then
                vs.TextMatrix(i, 2) = rs1!TypeofProduct + " (" + rs1!rulling + ")" + str(rs1!NoOfPages) + " " + rs1!ProductQuality
            End If
       
           vs.TextMatrix(i, 3) = rs(3)
           i = i + 1
           K = K + 1
           rs.MoveNext
           
     Wend
     
     cmdEdit_4.Enabled = True
     
     
  End If
  
  Total

End Sub

Private Sub cmdSave_2_Click()

If rs.State = 1 Then rs.Close
rs.Open "select PCode from ProductReceipt where " & stringyear & " and  recno=" & txtRecNo & "", CON, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then
   CON.Execute "delete from ProductReceipt where " & stringyear & " and  recno=" & txtRecNo & ""
Else
   MaxRecNo
End If


If MsgBox("Want To Save ?", vbInformation + vbYesNo) = vbYes Then
   


For i = 1 To vs.Rows - 1
   
    If vs.TextMatrix(i, 1) <> "" Then
       
       CON.Execute "insert into ProductReceipt(" & _
          "[RecNo],[RecDate]" & _
          ",[PCode]" & _
          ",[QUANTITY]" & _
          ",[fyear]" & _
          ",[Createdby]" & _
          ",[createdon]" & _
          ",[updatedby]" & _
          ",[updatedon]" & _
          ",[setupid]) values(" & txtRecNo.Text & ",'" & Format(RecDate.Text, "MM/dd/yyyy") & "'," & _
          "'" & vs.TextMatrix(i, 1) & "'," & Val(vs.TextMatrix(i, 3)) & ",'" & main.session & "'," & _
          "'" & main.username & "','" & Format(Date, "MM/dd/yyyy") & "','" & main.username & "','" & Format(Date, "MM/dd/yyyy") & "','" & main.setupid & "')"
          
    End If
      
Next
      
      
cmdSave_2.Enabled = False
Call cmdAdd_1_Click
txtRem.SetFocus


      
End If

End Sub

Private Sub cmdSearch_Click()
popuplist10 "select  [RecNo],[RecDate],[fyear] FROM [ProductReceipt] where " & stringyear & " group by [RecNo],[RecDate],[fyear]", CON
End Sub

Private Sub cmdSearch_GotFocus()
If PopUpValue1 <> "" Then
   txtRecNo = PopUpValue1
   searchData
   PopUpValue1 = ""
End If

End Sub

Private Sub Form_Load()

setwidth

MaxRecNo

RecDate.Text = Format(Date, "dd/MM/yyyy")

End Sub
Sub MaxRecNo()
    If rs.State = 1 Then rs.Close
    rs.Open "select max(RecNo) from [ProductReceipt] where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly
    If Not IsNull(rs(0)) Then
       txtRecNo = rs(0) + 1
    Else
       txtRecNo = 1
    End If
    
    
End Sub

Sub setwidth()
    vs.Clear
    vs.FormatString = "SR.|P.Code|Description|Quantity"
    
    vs.ColWidth(0) = 500
    vs.ColWidth(1) = 1400
    vs.ColWidth(2) = 6000
    vs.ColWidth(3) = 1500
    
    
    
End Sub
Sub Total()
On Error Resume Next

txtTotal.Text = 0
For i = 1 To vs.Rows - 1
If vs.TextMatrix(i, 3) <> "" Then
txtTotal.Text = (Val(txtTotal.Text) + vs.TextMatrix(i, 3))
End If
Next

End Sub

Private Sub RecDate_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub txtRecNo_GotFocus()
HIT


End Sub

Private Sub txtRecNo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    searchData
    SendKeys "{tab}"
 End If
End Sub
Private Sub txtRem_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  SendKeys "{tab}"
  vs.Row = 1
  
 End If
End Sub

Private Sub VS_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 46 Then
   CON.Execute "delete from ProductReceipt where " & stringyear & " and  RecNo=" & txtRecNo & " and pcode='" & vs.TextMatrix(vs.RowSel, 1) & "'"
   vs.RemoveItem (vs.RowSel)
End If

End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    
    If vs.Col = 1 Then
    
    If vs.TextMatrix(vs.RowSel, 1) <> "" Then
    
    
    If rs.State = 1 Then rs.Close
        rs.Open "select ProductQuality,TypeofProduct,rulling,rate,NoofPages from copymaster " & _
            "where " & stringyear & " and  bookno='" & vs.TextMatrix(vs.RowSel, 1) & "'", CON
            If rs.EOF = False Then
                vs.TextMatrix(vs.RowSel, 2) = rs!TypeofProduct + " (" + rs!rulling + ")" + str(rs!NoOfPages) + " " + rs!ProductQuality
            Else
                Exit Sub
            End If

           SendKeys "{Right}"
           SendKeys "{Right}"
     End If
    
    
    ElseIf vs.Col = 3 Then
    
    If vs.TextMatrix(vs.RowSel, 3) <> "" Then
       SendKeys "{home}"
       SendKeys "{down}"
       Total
    End If
    
    End If
    
    End If
End Sub
