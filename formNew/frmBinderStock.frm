VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBinderStock 
   Caption         =   "Binder Stock Summary"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   13836
   Icon            =   "frmBinderStock.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   13836
   Begin VB.TextBox txtbk 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   14
      Top             =   1008
      Width           =   1884
   End
   Begin VB.CommandButton Command1_op 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Book Closing Tranfer"
      Height          =   510
      Left            =   9135
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   135
      Width           =   4380
   End
   Begin VB.CommandButton cmdupdate_ 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update From Chitra Soft"
      Height          =   645
      Left            =   13644
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   684
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export to Excel"
      Height          =   645
      Left            =   10704
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   675
      Width           =   1416
   End
   Begin VB.ComboBox cbostockType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      ItemData        =   "frmBinderStock.frx":000C
      Left            =   6885
      List            =   "frmBinderStock.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   672
      Width           =   2175
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View"
      Height          =   645
      Left            =   9135
      Picture         =   "frmBinderStock.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   675
      Width           =   1536
   End
   Begin VB.CommandButton cmdExit1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   645
      Left            =   12144
      Picture         =   "frmBinderStock.frx":0C1C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   675
      Width           =   1380
   End
   Begin VB.OptionButton Option2_itc 
      Caption         =   "ITC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1575
      TabIndex        =   4
      Top             =   264
      Value           =   -1  'True
      Width           =   780
   End
   Begin VB.OptionButton Option1_new 
      Caption         =   "New Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2475
      TabIndex        =   3
      Top             =   264
      Width           =   1410
   End
   Begin VB.ComboBox cboBinder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1575
      TabIndex        =   2
      Top             =   672
      Width           =   5280
   End
   Begin MSComCtl2.DTPicker dateAson 
      Height          =   336
      Left            =   5472
      TabIndex        =   0
      Top             =   288
      Visible         =   0   'False
      Width           =   1392
      _ExtentX        =   2455
      _ExtentY        =   593
      _Version        =   393216
      Format          =   498401281
      CurrentDate     =   39795
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7356
      Left            =   132
      TabIndex        =   7
      Top             =   1356
      Width           =   13416
      _cx             =   23664
      _cy             =   12975
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
      Rows            =   200
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
   Begin VB.Label Label1_bk 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name (f2 for search book)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   144
      TabIndex        =   15
      Top             =   1008
      Width           =   2820
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key To Delete A Invoive Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   180
      TabIndex        =   11
      Top             =   8775
      Width           =   2955
   End
   Begin VB.Label lblstock 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   336
      Index           =   1
      Left            =   6936
      TabIndex        =   9
      Top             =   312
      Width           =   1500
   End
   Begin VB.Label binderlbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Binder Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   180
      TabIndex        =   1
      Top             =   684
      Width           =   1236
   End
End
Attribute VB_Name = "frmBinderStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim btype As String
Private Sub cboBinder_Click()

On Error GoTo aa1:


If bookOp = "y" Then

vs.Clear

If RS.State = 1 Then RS.close
RS.Open "SELECT BookCode,BookName,op FROM BinderBookOp  where " & _
" binder='" & cboBinder & "' and ITC_New='" & btype & "'    order by BookCode", con
For h1 = 1 To RS.RecordCount
   
   DoEvents
   
   If RS.EOF = False Then
   
    vs.rows = vs.rows + 1
    vs.TextMatrix(h1, 0) = RS(0)
    vs.TextMatrix(h1, 1) = RS(1)
    vs.TextMatrix(h1, 2) = RS(2)
    RS.MoveNext
   
   End If
Next

End If

vsCol_op

Exit Sub

aa1:

MsgBox "" & err.DESCRIPTION

End Sub

Private Sub cbostockType_Click()

If cbostockType.text = "Stock Summary" Then

txtbk.Visible = False
Label1_bk.Visible = False

Else

txtbk.Visible = True
Label1_bk.Visible = True

End If

End Sub

Private Sub cmdExit1_Click()
 Unload Me
End Sub
Private Sub cmdRepQty_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String


Screen.MousePointer = vbHourglass



If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double

Dim b1 As Boolean

b1 = False


c = 1
r = 1



row_ = 4
col_ = 1

 xl.Columns("A:H").ColumnWidth = 12
 J = 2
 xlSheet.Cells(1, 1).value = cboBinder.text
 xlSheet.Cells(2, 1).value = Format(Date, "dd/MM/yyyy")
 
 
 For I = 0 To vs.rows - 1
     For J = 0 To vs.Cols - 1
            xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
           col_ = col_ + 1
     Next
     row_ = row_ + 1
     col_ = 1
 Next
    
 Screen.MousePointer = vbDefault


End Sub

Private Sub cmdupdate__Click()
Dim con_chitra As New ADODB.Connection


''st1 = "\\Server\bookinvent"
''con_chitra.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + st1 + "\" + Trim("" & session) + "\data.mdb"
''con_chitra.Open

Dim db_ As String
Dim sqluser, sqlpass As String

sqluser = "chitradatabase"
sqlpass = "java.123"
db_ = "Database=chitraDNet_" & Right(databaseNew, 4)

Set con_chitra = New ADODB.Connection
    
serverNameNew_ = "WIN-FI4EQR95VL3\CHITRASQLserver"
    
con_chitra.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew_ & "; " & db_ & "; UID=" & sqluser & "; PWD=" & sqlpass
    
DoEvents
DoEvents
    
con_chitra.CursorLocation = adUseClient
If con_chitra.State = 1 Then con_chitra.close
con_chitra.Open



Screen.MousePointer = vbHourglass

con.Execute "delete from BookReceiveDet_chitra"

s10_ = "SELECT a.INVOICENO,b.invoicedate,'Sundry Debtors' as Genledger,b.Binder_name as subledger ,c.bookname as BOOKCODE,a.Gaddi as tbook,a.BkInGaddi as loosbook," & _
"a.LooseBk as TotalBook,a.TotalBk as NetBook,a.BOOKCODE as book_code,a.Remarks,b.Fyear,b.setupid,c.GROUPCODE " & _
"FROM BinderBkReceiveDet as a inner join BinderReceiveBkQry as b on (a.INVOICENO = b.INVOICENO) inner join books as c on (a.BOOKCODE = c.BOOKCODE)"

Set rs1 = New ADODB.Recordset
rs1.Open "select INVOICENO,INVOICEDATE,Genledger,SUBLEDGER,bookcode,tbook,loosbook,TotalBook,netbook,Book_Code,remarks,fyear,setupid from BookReceiveDet_chitra", con, adOpenDynamic, adLockOptimistic

Set RS = New ADODB.Recordset
RS.Open s10_, con_chitra
While RS.EOF = False
 
 rs1.AddNew
 rs1!invoiceNo = RS!invoiceNo
 rs1!invoiceDate = RS!invoiceDate
 rs1!Genledger = RS!Genledger
 rs1!subledger = RS!subledger
 rs1!Bookcode = RS!Bookcode
 rs1!tbook = RS!tbook
 rs1!loosbook = RS!loosbook
 rs1!TotalBook = RS!TotalBook
 rs1!netbook = RS!netbook
 rs1!Book_Code = RS!Book_Code
 rs1!remarks = RS!remarks
 rs1!fyear = session
 rs1!setupid = 1
 rs1!groupcode = RS!groupcode
 
 rs1.update
RS.MoveNext
Wend

con_chitra.close

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdView_Click()



On Error GoTo aaa1:



''rs_p.Open "select distinct bookno,book from BookMaster", con, adOpenDynamic, adLockOptimistic
'aa_ = Mid(cboBinder.Text, 1, 5)
'If aa_ = "ILYAS" Then
'    aa_ = Mid(cboBinder.Text, 1, 2)
'    con.Execute "update BookReceiveDet_chitra set subledger ='" & cboBinder.Text & "' where subledger like '" & aa_ & "%'"
'ElseIf cboBinder.Text = "AZAM BINDER" Then
'    con.Execute "update BookReceiveDet_chitra set subledger ='AZAM BINDER' where subledger = 'M.A.BOOK BINDING HOUSE'"
'ElseIf cboBinder.Text = "SALEEM BINDER" Then
'    con.Execute "update BookReceiveDet_chitra set subledger ='SALEEM BINDER' where subledger = 'SALEEM & BROTHERS'"
'ElseIf cboBinder.Text = "ZAHEER BINDER" Then
'    con.Execute "update BookReceiveDet_chitra set subledger ='ZAHEER BINDER' where subledger = 'ZAHEER BOOK BINDING HOUSE'"
'ElseIf cboBinder.Text = "RIYAAZ BINDER" Then
'    con.Execute "update BookReceiveDet_chitra set subledger ='RIYAAZ BINDER' where subledger = 'N.S.BOOK BINDING HOUSE'"
'End If


If Option2_itc.value = True Then
 
 vs.rows = 1
 
  
 If cbostockType.text = "Stock Summary" Then
        
        vs.Cols = 6
        
        If RS.State = 1 Then RS.close
        RS.Open "SELECT BookCode,b.book,sum(IssueQty),sum(ReceiveQty),sum(op) FROM stock_ITCBinderStockQry " & _
        " left outer join BookMaster as b on (b.bookno=bookcode)  " & _
        " where convert(smalldatetime,invoiceDate,103)>= convert(smalldatetime,'" & dateAson & "',103) and " & _
        " subledger='" & cboBinder & "' group by b.book,Bookcode order by Bookcode", con
        For h1 = 1 To RS.RecordCount
           
           
           DoEvents
           DoEvents
           
           vs.rows = vs.rows + 1
           vs.TextMatrix(h1, 0) = RS(0)
           
           If Not IsNull(RS(1)) Then
           vs.TextMatrix(h1, 1) = RS(1)
           End If
           
           vs.TextMatrix(h1, 2) = RS(4)
           vs.TextMatrix(h1, 3) = RS(2)
           vs.TextMatrix(h1, 4) = RS(3)
           vs.TextMatrix(h1, 5) = ((RS(2) - RS(3)) + RS(4))

           
           
           DoEvents
           DoEvents
        
           RS.MoveNext
        Next
  
 Else
        vs.Cols = 7
 
        balQty = 0
        
        If RS.State = 1 Then RS.close
        
        If txtbk.text = "" Then
        
        RS.Open "SELECT BookCode,b.book,invoiceno,invoiceDate,IssueQty,ReceiveQty,op FROM stock_ITCBinderStockQry" & _
        " left outer join BookMaster as b on (b.bookno=bookcode)  " & _
        " where convert(smalldatetime,invoiceDate,103)>= convert(smalldatetime,'" & dateAson & "',103) and " & _
        " subledger='" & cboBinder & "' order by Bookcode,invoiceDate,invoiceno", con
        Else
        
           RS.Open "SELECT BookCode,b.book,invoiceno,invoiceDate,IssueQty,ReceiveQty,op FROM stock_ITCBinderStockQry" & _
        " left outer join BookMaster as b on (b.bookno=bookcode)  " & _
        " where convert(smalldatetime,invoiceDate,103)>= convert(smalldatetime,'" & dateAson & "',103) and " & _
        " subledger='" & cboBinder & "' and bookcode='" & txtbk.text & "' order by Bookcode,invoiceDate,invoiceno", con
    
        
        End If
        For h1 = 1 To RS.RecordCount
           
           DoEvents
           DoEvents
           
           If RS.EOF = False Then
           
           
           vs.rows = vs.rows + 1
           vs.TextMatrix(h1, 0) = RS(0)
           
           If Not IsNull(RS(1)) Then
              vs.TextMatrix(h1, 1) = RS(1)
           End If
           
           vs.TextMatrix(h1, 2) = RS!invoiceNo & " - " & RS!invoiceDate
           
           vs.TextMatrix(h1, 3) = RS(6) '' op
           
           vs.TextMatrix(h1, 4) = RS(4)
           vs.TextMatrix(h1, 5) = RS(5)
           
           vs.TextMatrix(h1, 6) = ((RS(4) - RS(5)) + RS(6))
           
           
           If h1 > 1 Then
              If bcode = RS(0) Then
                 vs.TextMatrix(h1, 6) = balQty + Val(vs.TextMatrix(h1, 6))
              Else
                 balQty = 0
              End If
           End If
           
           
           balQty = balQty + ((RS(4) - RS(5)) + RS(6))
           bcode = RS(0)
           
           
           DoEvents
           DoEvents
           
        
           RS.MoveNext
           
           End If
           
        Next
 
  
 End If

End If

'==================================================================
'==================================================================

If Option2_itc.value = False Then

 vs.rows = 1
 
  
 If cbostockType.text = "Stock Summary" Then
        
        vs.Cols = 6

        If RS.State = 1 Then RS.close
        RS.Open "SELECT a.BookCode,b.book,sum(a.IssueQty),sum(a.ReceiveQty),sum(a.op) FROM Stock_NewBinderStockQry as a " & _
        " left outer join BookMaster as b on (b.bookno=a.bookcode)  " & _
        " where convert(smalldatetime,a.invoiceDate,103)>= convert(smalldatetime,'" & dateAson & "',103) and " & _
        " (a.subledger='" & cboBinder & "' or a.linkto='" & cboBinder & "') group by a.Bookcode,b.book order by a.Bookcode", con
        For h1 = 1 To RS.RecordCount
           
           DoEvents
           DoEvents
           
           If RS.EOF = False Then
               vs.rows = vs.rows + 1
               vs.TextMatrix(h1, 0) = RS(0)
               If Not IsNull(RS(1)) Then
                 vs.TextMatrix(h1, 1) = RS(1)
               End If
               
               vs.TextMatrix(h1, 2) = RS(4)
              
               vs.TextMatrix(h1, 3) = RS(2)
               vs.TextMatrix(h1, 4) = RS(3)
               vs.TextMatrix(h1, 5) = ((RS(2) - RS(3)) + RS(4))
               
               DoEvents
               DoEvents
            
               RS.MoveNext
           End If
           
        Next
  
 Else
        vs.Cols = 7
        balQty = 0
        
        If RS.State = 1 Then RS.close
        
        If txtbk.text = "" Then
        
        RS.Open "SELECT BookCode,b.book,invoiceno,invoiceDate,IssueQty,ReceiveQty,op FROM Stock_NewBinderStockQry as a " & _
        " left outer join BookMaster as b on (b.bookno=a.bookcode)  " & _
        " where convert(smalldatetime,invoiceDate,103)>= convert(smalldatetime,'" & dateAson & "',103) and " & _
        " (subledger='" & cboBinder & "' or linkto='" & cboBinder & "') order by Bookcode,invoiceDate,invoiceno", con
        
        Else
        
        RS.Open "SELECT BookCode,b.book,invoiceno,invoiceDate,IssueQty,ReceiveQty,op FROM Stock_NewBinderStockQry as a " & _
        " left outer join BookMaster as b on (b.bookno=a.bookcode)  " & _
        " where convert(smalldatetime,invoiceDate,103)>= convert(smalldatetime,'" & dateAson & "',103) and " & _
        " (subledger='" & cboBinder & "' or linkto='" & cboBinder & "') and bookcode='" & txtbk.text & "' order by Bookcode,invoiceDate,invoiceno", con
        
        End If
        
        
        For h1 = 1 To RS.RecordCount
           
           DoEvents
           DoEvents
           
           vs.rows = vs.rows + 1
           vs.TextMatrix(h1, 0) = RS(0)
           If Not IsNull(RS(1)) Then
           vs.TextMatrix(h1, 1) = RS(1)
           End If
           
           vs.TextMatrix(h1, 2) = RS!invoiceNo & " - " & RS!invoiceDate
           
           vs.TextMatrix(h1, 3) = RS(6) '' op
           vs.TextMatrix(h1, 4) = RS(4)
           vs.TextMatrix(h1, 5) = RS(5)
           
           
           vs.TextMatrix(h1, 6) = ((RS(4) - RS(5)) + RS(6))
           
           
           If h1 > 1 Then
              If bcode = RS(0) Then
                 vs.TextMatrix(h1, 6) = balQty + Val(vs.TextMatrix(h1, 6))
              Else
                 balQty = 0
              End If
           End If
           
           
           balQty = balQty + ((RS(4) - RS(5)) + RS(6))
           bcode = RS(0)
           
           
           
           
           
           DoEvents
           DoEvents
        
           RS.MoveNext
        Next
 
  
 End If

End If





vsColSet

If bookOp = "y" Then
vsCol_op
End If

Exit Sub
aaa1:

MsgBox "" & err.DESCRIPTION

End Sub
Sub vsCol_op()
   
   vs.Cols = 3
   
   vs.FormatString = "**BookCode**|BookName|Opening"
    
   vs.ColWidth(0) = 1300
   vs.ColWidth(1) = 6000
   vs.ColWidth(2) = 1200
   
   
   'vs.ColHidden(3) = True
   'vs.ColHidden(4) = True
   'vs.ColHidden(5) = True
End Sub

Sub vsColSet()

If cbostockType.text = "Stock Summary" Then
   vs.Cols = 6
   
   vs.FormatString = "**BookCode**|BookName|Opening|Qty Issued|Qty Received|**Balance**"
    
   vs.ColWidth(0) = 1300
   vs.ColWidth(1) = 6000
   vs.ColWidth(2) = 1200
   vs.ColWidth(2) = 1200
   vs.ColWidth(3) = 1200
   vs.ColWidth(4) = 1200
   
   
Else

   vs.Cols = 7
   vs.FormatString = "**BookCode**|BookName|VNo. And Date|Opening|Qty Issued|Qty Received|**Balance**"
   vs.ColWidth(0) = 1300
   vs.ColWidth(1) = 4500
   vs.ColWidth(2) = 1700
   vs.ColWidth(3) = 1200
   vs.ColWidth(4) = 1200
   vs.ColWidth(5) = 1200
   vs.ColWidth(6) = 1200
   
    
End If

End Sub

Private Sub Command1_op_Click()

If LCase(UserName) = "admin" Then

   Dim CON_next As New ADODB.Connection
   Dim db_ As String
   Dim a1 As Integer
   

   
   a1 = Int(Right(databaseNew, 4)) + 101
   
   db_ = "Database=chitraData_" & a1
   Set CON_next = New ADODB.Connection
   CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
    
   DoEvents
   DoEvents
    
    CON_next.CursorLocation = adUseClient
    If CON_next.State = 1 Then CON_next.close
    CON_next.Open
     
      
  If MsgBox("want to transfer ?", vbQuestion + vbYesNo) = vbYes Then
   
     
     CON_next.Execute "delete from BinderBookOp"
     
     If RS.State = 1 Then RS.close
     RS.Open "SELECT BookCode ,sum((IssueQty-ReceiveQty)+OP) as Op,'ITC' as ITC_New,subledger as Binder,bookname FROM Stock_ITCBinderStockQry group by BookCode,subledger,bookname", con
     While RS.EOF = False
     
     CON_next.Execute "insert into BinderBookOp(BOOKCODE,op,ITC_New,binder,opdate,BookName)" & _
     " values('" & RS!Bookcode & "'," & RS(1) & ",'" & RS(2) & "','" & RS(3) & "','" & fromDate_setup & "','" & RS(4) & "')"
     
     RS.MoveNext
     Wend
   
     If RS.State = 1 Then RS.close
     RS.Open "SELECT BookCode ,sum((IssueQty-ReceiveQty)+OP) as Op,'New' as ITC_New,subledger as Binder,bookname FROM stock_NewBinderStockQry group by BookCode,subledger,bookname", con
     While RS.EOF = False
     
     CON_next.Execute "insert into BinderBookOp(BOOKCODE,op,ITC_New,binder,opdate,BookName)" & _
     " values('" & RS!Bookcode & "'," & RS(1) & ",'" & RS(2) & "','" & RS(3) & "','" & fromDate_setup & "','" & RS(4) & "')"
     
     RS.MoveNext
     Wend
   
   
  End If
  
Else
  
  MsgBox ("You are not Authorized ...")

End If


End Sub

Private Sub Form_Load()

Screen.MousePointer = vbHourglass


If bookOp = "n" Then
   Me.Caption = "Binder Stock Summary"
   cbostockType.Visible = True
   lblstock(1).Visible = True
Else
   Me.Caption = "Book Opening"
   cbostockType.Visible = False
   lblstock(1).Visible = False

End If


If LCase(UserName) = "admin" Then
   Command1_op.Visible = True
Else
   Command1_op.Visible = False
End If


con.Execute "update a set a.linkto=b.linkto FROM OrderPrint_Det as a " & _
"inner join Godownmaster as b on (a.Binder=b.Godwn)"


dateAson.value = Format(Date, "dd/MM/yyyy")

If session = "2020-21" Then
   dateAson = "25/06/2020"
Else
   dateAson = from_date
End If

Option2_itc_Click

Me.top = 100
Me.Left = 100

Me.Width = 15000
Me.Height = 9650



cbostockType.ListIndex = 0

'cbofirm.Clear
'If RS.State = 1 Then RS.close
'RS.Open "select FirmName from FirmMaster order by firmname", con, adOpenStatic, adLockReadOnly
'While RS.EOF = False
'  cbofirm.AddItem RS(0)
'  RS.MoveNext
'Wend

cboBinder.Clear
If RS.State = 1 Then RS.close
RS.Open "select Godwn from Godownmaster where  Len(Godwn) > 5 order by Godwn", con, adOpenStatic, adLockReadOnly
While RS.EOF = False
  cboBinder.AddItem RS(0)
  RS.MoveNext
Wend




vsCol_op

Screen.MousePointer = vbDefault

End Sub

Private Sub Option1_new_Click()
   If Option2_itc.value = True Then
      btype = "ITC"
   Else
      btype = "New"
   End If

End Sub
Private Sub Option2_itc_Click()
   If Option2_itc.value = True Then
      btype = "ITC"
   Else
      btype = "New"
   End If

End Sub


Private Sub txtbk_GotFocus()
If PopUpValue1 <> "" Then
    
   txtbk.text = PopUpValue1
    
   PopUpValue1 = ""
   PopUpValue2 = ""
   
End If

End Sub

Private Sub txtbk_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then

   searchType = "books"
   popuplist10 "SELECT distinct [BookNo] as BookCode,[book] as BookName  FROM BookMaster where firmname='BLUEPRINT EDUCATION' order by BookNo", con

End If

End Sub
Private Sub vs_DblClick()


Dim b1 As Boolean
If Val(vs.TextMatrix(vs.RowSel, 4)) > 0 Then
   b1 = True
   showvoucher vs.TextMatrix(vs.RowSel, 2), b1
Else
   b1 = False
   showvoucher vs.TextMatrix(vs.RowSel, 2), b1
End If



End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 115 Then
   
If bookOp = "y" Then

  If MsgBox("Want to delete ?", vbYesNo + vbQuestion) = vbYes Then
     If cboBinder.text <> "" Then
        con.Execute "delete BinderBookOp where bookcode='" & vs.TextMatrix(vs.RowSel, 0) & "' and ITC_New='" & btype & "' and Binder='" & cboBinder.text & "'"
        vs.RemoveItem (vs.RowSel)
     End If
     
  End If
  
End If

End If


If KeyCode = 13 Then
   Dim b1 As Boolean
   If Val(vs.TextMatrix(vs.RowSel, 4)) > 0 Then
      b1 = True
      showvoucher vs.TextMatrix(vs.RowSel, 2), b1
   Else
      b1 = False
      showvoucher vs.TextMatrix(vs.RowSel, 2), b1
   End If
End If



End Sub
Sub showvoucher(inv As String, issue_rec As Boolean)
   
   Screen.MousePointer = vbHourglass
   
   Dim inv_ As String
   
   inv_ = Mid(inv, 1, InStr(inv, "-") - 1)
   inviceNo = Trim(inv_)
   
   If issue_rec = True Then
      frmIssue.Show
   Else
      frmBinderRecChallan.Show
   End If
   
   Screen.MousePointer = vbDefault
   
End Sub


Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   
If bookOp = "y" Then

   
'   If Option2_itc.value = True Then
'      btype = "ITC"
'   Else
'      btype = "New"
'   End If
   
   
   If vs.Col = 0 Then
       
      If RS.State = 1 Then RS.close
      RS.Open "select * from books where bookcode='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
      If RS.EOF = False Then
         vs.TextMatrix(vs.RowSel, 0) = UCase(vs.TextMatrix(vs.RowSel, 0))
         vs.TextMatrix(vs.RowSel, 1) = RS!Bookname
         sendkeys "{right}"
         sendkeys "{right}"

      End If
   
   End If
   
   If vs.Col = 2 Then
   
   
   
      If RS.State = 1 Then RS.close
      RS.Open "select * from BinderBookOp where (bookcode='" & vs.TextMatrix(vs.RowSel, 0) & "' and ITC_New='" & btype & "' and binder='" & cboBinder.text & "')", con
      If RS.EOF = True Then
         con.Execute "insert into BinderBookOp(BOOKCODE,bookName,OP,ITC_New,Binder,opdate) values('" & vs.TextMatrix(vs.RowSel, 0) & "','" & vs.TextMatrix(vs.RowSel, 1) & "','" & IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) & "','" & btype & "','" & cboBinder.text & "','" & Format(dateAson, "MM/dd/yyyy") & "')"
      Else
         con.Execute "update BinderBookOp set op='" & IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) & "' where (bookcode='" & vs.TextMatrix(vs.RowSel, 0) & "' and ITC_New='" & btype & "' and binder='" & cboBinder.text & "')"
      End If
      
      sendkeys "{home}"
      sendkeys "{down}"
   End If
  
   
   
End If
End If

End Sub

Private Sub vs_SelChange()
If bookOp = "y" Then
   If vs.Col = 2 Then
      vs.Editable = flexEDKbdMouse
   End If
Else
   If vs.Col = 2 Then
      vs.Editable = flexEDNone
   End If
End If
End Sub
