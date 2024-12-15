VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIssueRecAuth 
   Caption         =   "Authentication Form"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   Icon            =   "frmIssueRecAuth.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   2460
      TabIndex        =   10
      Top             =   600
      Width           =   7245
      Begin VB.OptionButton autho 
         Caption         =   "Authorized"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   255
         Width           =   1335
      End
      Begin VB.OptionButton Unautho 
         Caption         =   "Un Authorized"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1545
         TabIndex        =   12
         Top             =   195
         Width           =   1620
      End
      Begin VB.OptionButton All 
         Caption         =   "All"
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
         Left            =   3300
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdset 
      BackColor       =   &H00FFFFC0&
      Caption         =   "S&ave"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   780
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   2460
      TabIndex        =   1
      Top             =   0
      Width           =   7245
      Begin VB.OptionButton Option1_binderRec 
         Caption         =   "Binder Book Rec."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         TabIndex        =   14
         Top             =   240
         Width           =   1800
      End
      Begin VB.OptionButton BookStockTrans 
         Caption         =   "Book Stock Transfer"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   4
         Top             =   210
         Width           =   1980
      End
      Begin VB.OptionButton BookReceive 
         Caption         =   "Book Receive List"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1650
         TabIndex        =   3
         Top             =   210
         Width           =   1800
      End
      Begin VB.OptionButton BookIssue 
         Caption         =   "Book Issue List"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   60
         TabIndex        =   2
         Top             =   210
         Width           =   1620
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7500
      Left            =   60
      TabIndex        =   0
      Top             =   1320
      Width           =   9600
      _cx             =   16933
      _cy             =   13229
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777209
      ForeColor       =   16711680
      BackColorFixed  =   16777173
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777166
      BackColorAlternate=   16777209
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIssueRecAuth.frx":000C
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
      ExplorerBar     =   7
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
   Begin MSComCtl2.DTPicker toDate 
      Height          =   315
      Left            =   855
      TabIndex        =   6
      Top             =   390
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   71434241
      CurrentDate     =   38845
   End
   Begin MSComCtl2.DTPicker FromDate 
      Height          =   300
      Left            =   840
      TabIndex        =   7
      Top             =   60
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   529
      _Version        =   393216
      Format          =   71434241
      CurrentDate     =   38845
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   450
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date "
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frmIssueRecAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub All_Click()
If All.value = True Then
   fillGrid
End If
End Sub

Private Sub autho_Click()
If autho.value = True Then
   fillGrid
End If
End Sub

Private Sub BookIssue_Click()
If BookIssue.value = True Then
   fillGrid
End If
End Sub

Private Sub BookReceive_Click()
If BookReceive.value = True Then
   fillGrid
End If
End Sub

Private Sub BookStockTrans_Click()

If BookStockTrans.value = True Then
   fillGrid
End If

End Sub

Private Sub cmdset_Click()

If RS.State = 1 Then RS.close
RS.Open "select * from pass where pass='" & cp & "'", con
If RS.EOF = True Then
   MsgBox "Enter Valid Password !!", vbInformation
Exit Sub
End If
    
saveData

End Sub
Sub saveData()
   
On Error GoTo ss:
   
   If MsgBox("Want to Set ?", vbQuestion + vbYesNo) = vbYes Then
        
   Screen.MousePointer = vbHourglass
   
   Dim din As Integer
         
   If BookIssue.value = True Then
        
        For J = 1 To vs.Rows - 1
          
        If vs.TextMatrix(J, 5) <> "" Then
          
          If vs.TextMatrix(J, 5) = True Then
             If vs.TextMatrix(J, 5) = True Then din = 1
             con.Execute "update BookStock set bAuthorized=" & din & " where " & stringyear & " and EntryNo=" & vs.TextMatrix(J, 1) & ""
            Else
             If vs.TextMatrix(J, 5) = False Then din = 0
             con.Execute "update BookStock set bAuthorized=" & din & " where " & stringyear & " and EntryNo=" & vs.TextMatrix(J, 1) & ""
          End If
        
        End If
          
          
        Next
    
 ElseIf BookReceive.value = True Then
        
        For J = 1 To vs.Rows - 1
          
        If vs.TextMatrix(J, 5) <> "" Then
          
          If vs.TextMatrix(J, 5) = True Then
             If vs.TextMatrix(J, 5) = True Then din = 1
             con.Execute "update BookStock set bAuthorized=" & din & " where " & stringyear & " and EntryNo=" & vs.TextMatrix(J, 1) & ""
            Else
             If vs.TextMatrix(J, 5) = False Then din = 0
             con.Execute "update BookStock set bAuthorized=" & din & " where " & stringyear & " and EntryNo=" & vs.TextMatrix(J, 1) & ""
          End If
        
        End If
          
          
        Next
        
 ElseIf BookStockTrans.value = True Then
        
        For J = 1 To vs.Rows - 1
          
        If vs.TextMatrix(J, 5) <> "" Then
          
          If vs.TextMatrix(J, 5) = True Then
             If vs.TextMatrix(J, 5) = True Then din = 1
             con.Execute "update BookStock set bAuthorized=" & din & " where " & stringyear & " and EntryNo=" & vs.TextMatrix(J, 1) & ""
            Else
             If vs.TextMatrix(J, 5) = False Then din = 0
             con.Execute "update BookStock set bAuthorized=" & din & " where " & stringyear & " and EntryNo=" & vs.TextMatrix(J, 1) & ""
          End If
        
        End If
          
          
        Next
   
 ElseIf Option1_binderRec.value = True Then
        
        For J = 1 To vs.Rows - 1
          
        If vs.TextMatrix(J, 5) <> "" Then
          
          If vs.TextMatrix(J, 5) = True Then
             If vs.TextMatrix(J, 5) = True Then din = 1
             con.Execute "update BinderBkReceive set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
            Else
             If vs.TextMatrix(J, 5) = False Then din = 0
             con.Execute "update BinderBkReceive set bAuthorized=" & din & " where " & stringyear & " and INVOICENO=" & vs.TextMatrix(J, 1) & ""
          End If
        
        End If
          
          
        Next
   
   
  End If
   
   
End If
   

Screen.MousePointer = vbDefault

Exit Sub
ss:
MsgBox "" & "Connection Not Created Properly !", vbInformation
Screen.MousePointer = vbDefault


End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
   Unload Me
End If

End Sub
Private Sub Form_Load()
 Me.FromDate.value = from_date
 Me.toDate.value = to_date
 
 Me.Top = 800
 Me.Left = 200
 
 Me.Width = 9810
 Me.Height = 9300
 
End Sub
Sub vsIni()
   vs.FormatString = "SNo|Entry No|Date|Book Detils|>Quantity|^Entry-Authorization"
   vs.ColWidth(0) = 900
   vs.ColWidth(1) = 1500
   vs.ColWidth(2) = 1100
   vs.ColWidth(3) = 3200
   vs.ColWidth(4) = 1300
   vs.ColWidth(5) = 1200
   
End Sub
Sub fillGrid()
      
Dim datecondition As String
      
Screen.MousePointer = vbHourglass
vs.Clear

datecondition = "convert(smalldatetime,Dates,103)>=convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,Dates,103)<=convert(smalldatetime,'" & toDate.value & "',103)"
      
If BookIssue.value = True Then
      
      If RS.State = 1 Then RS.close
      If All.value = True Then
         RS.Open "SELECT EntryNo,Dates,Category,sum(Qty),BAuthorized FROM BookStock where " & stringyear & " and " & datecondition & " and Issue_Receive='Issue' and Category<>'Transfer' group BY EntryNo,Dates,Category,BAuthorized", con
      ElseIf autho.value = True Then
         RS.Open "SELECT EntryNo,Dates,Category,sum(Qty),BAuthorized FROM BookStock where " & stringyear & " and " & datecondition & " and Issue_Receive='Issue' and Category<>'Transfer' and bAuthorized=1 group BY EntryNo,Dates,Category,BAuthorized", con
      Else
         RS.Open "SELECT EntryNo,Dates,Category,sum(Qty),BAuthorized FROM BookStock where " & stringyear & " and " & datecondition & " and Issue_Receive='Issue' and Category<>'Transfer' and bAuthorized=0 group BY EntryNo,Dates,Category,BAuthorized", con
      End If
      
      
      If RS.EOF = False Then
        vs.Rows = RS.RecordCount + 1
        For I = 1 To vs.Rows - 1
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           RS.MoveNext
        Next
      Else
           vs.Clear
           vs.Rows = 2
      End If
      
ElseIf BookReceive.value = True Then

      If RS.State = 1 Then RS.close
      If All.value = True Then
         RS.Open "SELECT EntryNo,Dates,Category,sum(Qty),BAuthorized FROM BookStock where " & stringyear & " and " & datecondition & " and Issue_Receive='Receive' and Category<>'Transfer' group BY EntryNo,Dates,Category,BAuthorized", con
      ElseIf autho.value = True Then
         RS.Open "SELECT EntryNo,Dates,Category,sum(Qty),BAuthorized FROM BookStock where " & stringyear & " and " & datecondition & " and Issue_Receive='Receive' and Category<>'Transfer' and bAuthorized=1 group BY EntryNo,Dates,Category,BAuthorized", con
      Else
         RS.Open "SELECT EntryNo,Dates,Category,sum(Qty),BAuthorized FROM BookStock where " & stringyear & " and " & datecondition & " and Issue_Receive='Receive' and Category<>'Transfer' and bAuthorized=0 group BY EntryNo,Dates,Category,BAuthorized", con
      End If
      
      
      If RS.EOF = False Then
        vs.Rows = RS.RecordCount + 1
        For I = 1 To vs.Rows - 1
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           RS.MoveNext
        Next
      Else
           vs.Clear
           vs.Rows = 2
      End If


ElseIf BookStockTrans.value = True Then

      If RS.State = 1 Then RS.close
      If All.value = True Then
         RS.Open "SELECT EntryNo,Dates,Category,sum(Qty),BAuthorized FROM BookStock where " & stringyear & " and " & datecondition & " and Issue_Receive='Issue' and Category='Transfer' group BY EntryNo,Dates,Category,BAuthorized", con
      ElseIf autho.value = True Then
         RS.Open "SELECT EntryNo,Dates,Category,sum(Qty),BAuthorized FROM BookStock where " & stringyear & " and " & datecondition & " and Issue_Receive='Issue' and Category='Transfer' and bAuthorized=1 group BY EntryNo,Dates,Category,BAuthorized", con
      Else
         RS.Open "SELECT EntryNo,Dates,Category,sum(Qty),BAuthorized FROM BookStock where " & stringyear & " and " & datecondition & " and Issue_Receive='Issue' and Category='Transfer' and bAuthorized=0 group BY EntryNo,Dates,Category,BAuthorized", con
      End If
      
      
      If RS.EOF = False Then
        vs.Rows = RS.RecordCount + 1
        For I = 1 To vs.Rows - 1
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           vs.TextMatrix(I, 4) = RS.Fields(3).value
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           RS.MoveNext
        Next
      Else
           vs.Clear
           vs.Rows = 2
      End If

ElseIf Option1_binderRec.value = True Then
      
      Dim rs1 As New ADODB.Recordset

      'datecondition = "(INVOICEDATE>=datevalue('" & Format(FromDate.value, "dd/MM/yy") & "') and INVOICEDATE<=datevalue('" & Format(toDate.value, "dd/MM/yy") & "'))"
      datecondition = "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & FromDate.value & "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & toDate.value & "',103)"
      
      If RS.State = 1 Then RS.close
      If All.value = True Then
         RS.Open "SELECT INVOICENO,INVOICEDATE,SUBLEDGER,Remarks,BAuthorized FROM BinderBkReceive where  " & datecondition & "", con
      ElseIf autho.value = True Then
         RS.Open "SELECT INVOICENO,INVOICEDATE,SUBLEDGER,Remarks,BAuthorized FROM BinderBkReceive where  " & datecondition & " and bAuthorized=1 ", con
      Else
         RS.Open "SELECT INVOICENO,INVOICEDATE,SUBLEDGER,Remarks,BAuthorized FROM BinderBkReceive where  " & datecondition & " and bAuthorized=0 ", con
      End If
      
      
      If RS.EOF = False Then
        vs.Rows = RS.RecordCount + 1
        For I = 1 To vs.Rows - 1
           vs.TextMatrix(I, 0) = I
           vs.TextMatrix(I, 1) = RS.Fields(0).value
           vs.TextMatrix(I, 2) = RS.Fields(1).value
           vs.TextMatrix(I, 3) = RS.Fields(2).value
           If rs1.State = 1 Then rs1.close
           rs1.Open "select sum(NetBook) from BookReceiveDet where INVOICENO=" & RS.Fields(0).value & "", con
           If Not IsNull(rs1(0)) Then
           vs.TextMatrix(I, 4) = rs1.Fields(0).value
           End If
           vs.TextMatrix(I, 5) = RS.Fields(4).value & ""
           RS.MoveNext
        Next
      Else
           vs.Clear
           vs.Rows = 2
      End If

End If

      

'=================================================================


 vsIni
 Screen.MousePointer = vbDefault
 

End Sub

Private Sub Option1_binderRec_Click()

If Option1_binderRec.value = True Then
   fillGrid
End If


End Sub

Private Sub Unautho_Click()
If Unautho.value = True Then
   fillGrid
End If

End Sub
