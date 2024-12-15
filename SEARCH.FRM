VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form searchscreen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5400
   ClientLeft      =   2040
   ClientTop       =   2040
   ClientWidth     =   9225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Search.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3345
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   5900
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox textsearch 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   4830
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Search String"
      Height          =   225
      Left            =   510
      TabIndex        =   0
      Top             =   4830
      Width           =   1185
   End
End
Attribute VB_Name = "searchscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim tablabel As Integer
Dim masterlabel As String
Dim maxrow As Integer



Private Sub Form_KeyPress(KeyAscii As Integer)
Grid1.V
    If KeyAscii = 27 Then
        Unload Me
        If Trim(UCase(masterlabel)) = Trim(UCase("master")) Then
            master.Enabled = True
        Else
            If Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
                bookmaster.Enabled = True
            Else
                If Trim(UCase(masterlabel)) = Trim(UCase("invoice")) Then
                    INVOICE.Enabled = True
                End If
                If Trim(UCase(masterlabel)) = Trim(UCase("credititemnote")) Then
                    CRITNOTE.Enabled = True
                End If
                
                If Trim(UCase(masterlabel)) = Trim(UCase("countersale")) Then
                    countersale.Enabled = True
                End If
                
               If Trim(UCase(masterlabel)) = Trim(UCase("Debitnotefile")) Then
                    Debitnotefile.Show
              End If
                
              If Trim(UCase(masterlabel)) = Trim(UCase("Creditnotefile")) Then
                    Creditnotefile.Show
              End If
                
                
                
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Grid1.RowExpanded = True
    Grid1.row = 0
    Grid1.col = 0
    Set con = New ADODB.Connection
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
    con.Open
    Set rs = New ADODB.Recordset
       'grid1.Rows = 1 'adarsh
    Me.Left = 1000
    Me.Top = 600
    End Sub
Private Sub Form_Terminate()
    If Trim(UCase(masterlabel)) = Trim(UCase("master")) Then
        master.Enabled = True
    Else
        If Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
            bookmaster.Enabled = True
        Else
            If Trim(UCase(masterlabel)) = Trim(UCase("invoice")) Then
                INVOICE.Enabled = True
            End If
            
            If Trim(UCase(masterlabel)) = Trim(UCase("countersale")) Then
                countersale.Enabled = True
            End If

            
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Trim(UCase(masterlabel)) = Trim(UCase("master")) Then
        master.Enabled = True
    Else
        If Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
            bookmaster.Enabled = True
        Else
            If Trim(UCase(masterlabel)) = Trim(UCase("invoice")) Then
                INVOICE.Enabled = True
            Else
               If Trim(UCase(masterlabel)) = Trim(UCase("voucher")) Then
                    Voucherform.Enabled = True
               End If
                If Trim(UCase(masterlabel)) = Trim(UCase("CREDITITEMNOTE")) Then
                    CRITNOTE.Enabled = True
                End If
                
                If Trim(UCase(masterlabel)) = Trim(UCase("countersale")) Then
                    countersale.Enabled = True
                End If
                
                
                If Trim(UCase(masterlabel)) = Trim(UCase("CREDITnotefile")) Then
                    Creditnotefile.Enabled = True
                End If
                
                If Trim(UCase(masterlabel)) = Trim(UCase("Debitnotefile")) Then
                    Debitnotefile.Enabled = True
                End If
                
                
            End If
        End If
    End If
End Sub

Private Sub Grid1_DblClick()
'GRID1.Rows
If Grid1.Rows > 1 Then
    If Trim(UCase(masterlabel)) = UCase("master") Then
        For I = 0 To 5
            If master.SSTab1.Tab <> I Then
                master.SSTab1.TabEnabled(I) = False
            End If
        Next
    Else
        If Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
            For I = 0 To 3
                If bookmaster.SSTab1.Tab <> I Then
                    bookmaster.SSTab1.TabEnabled(I) = False
                End If
            Next
        End If
    End If
    
    
'by vk  'master.Commandmasteradd.Enabled = True
        'master.Commandmasteredit.Enabled = True
    If tablabel = 0 And Trim(UCase(masterlabel)) = UCase("master") Then
        
        master.Commandmasteradd.Enabled = True
        master.Commandmasteredit.Enabled = True
        
        Grid1.col = 0
        master.ComboSPECIALCATEGORY = Grid1.Text
        Grid1.col = 1
        master.Textglgeneralledgerdiscription = Grid1.Text
        master.Textfindgl.Text = Grid1.Text
        Grid1.col = 2
        If Grid1.Text = True Then
            master.GMASTERPL.Value = 1
        Else
            master.GMASTERPL.Value = 0
        End If
        Grid1.col = 3
        If Grid1.Text = True Then
            master.GMASTERBS.Value = 1
        Else
            master.GMASTERBS.Value = 0
        End If
        
        Grid1.col = 4
        If Grid1.Text = True Then
            master.GMASTERSL.Value = 1
        Else
            master.GMASTERSL.Value = 0
        End If
        Grid1.col = 5
        master.Textglyearopeningbalance = Trim(Grid1.Text)
        Grid1.col = 6
        If Grid1.Text = True Then
            master.Cashbankbook.Value = 1
        Else
            master.Cashbankbook.Value = 0
        End If
            Unload Me
            master.Enabled = True
        If master.Commandmasteradd.Enabled = False Then
            master.Commandmastersave.Enabled = True
        End If
         master.Commandmasteradd.Enabled = True
         master.Commandmasteredit.Enabled = True
        If master.Commandmasteradd.Enabled = True Then
             If master.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
               
                If X = 6 Then
                    master.gledger.Enabled = True
                    master.Textglyearopeningbalance.Enabled = True
                    master.GMASTERPL.Enabled = True
                    master.GMASTERBS.Enabled = True
                    master.GMASTERSL.Enabled = True
                    master.Cashbankbook.Enabled = True
                    master.Textglgeneralledgerdiscription.Enabled = False
                    master.ComboSPECIALCATEGORY.Enabled = True
                    master.Commandmastersave.Enabled = True
                    master.Commandmasteredit.Enabled = False
                    master.Commandmasteradd.Enabled = False
                    master.Commandmasterdelete.Enabled = True
                    master.Commandmasterabandon.Enabled = True
                  '  master.Textglgeneralledgerdiscription.Enabled = True
                    master.Textglyearopeningbalance.Enabled = True
                    master.ComboSPECIALCATEGORY.SetFocus
                Else
                                    
                    master.gledger.Enabled = False
                End If
                
                
                
                
                
            End If
        End If
    End If
'///////////////********************////////
'       SUB LEADGER SEARCH
'///////////////********************////////
    
    If tablabel = 1 And Trim(UCase(masterlabel)) = UCase("master") Then
        master.Commandmasteradd.Enabled = True
        master.Commandmasteredit.Enabled = True

        Grid1.col = 0
        master.Comboslgenledgerdiscription = Grid1.Text
        Grid1.col = 1
        master.Textslsubledgerdiscription = Grid1.Text
     
        master.TextFINDSUBLEADGER = Grid1.Text
        Grid1.col = 2
        master.Textsldiscriptionforinvoice = Grid1.Text
        Grid1.col = 3
        master.Textslyearopeningbalance = Grid1.Text
        Grid1.col = 4
        master.Combosldiscountcategory = Grid1.Text
        Grid1.col = 5
        master.Combosldistrictcode = Grid1.Text
        Grid1.col = 6
        master.Textsladdress1 = Grid1.Text
        Grid1.col = 7
        master.Textsladdress2 = Grid1.Text
        Grid1.col = 8
        master.Textsladdress3 = Grid1.Text
        Unload Me
        master.Enabled = True
        If master.Commandmasteradd.Enabled = False Then
            master.Commandmastersave.Enabled = True
        End If
        If master.Commandmasteradd.Enabled = True Then
            If master.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                    master.sledger.Enabled = True
                    master.Textslsubledgerdiscription.Enabled = True
                   'master.Comboslgenledgerdiscription.Enabled = True
                    master.Textsldiscriptionforinvoice.Enabled = True
                    master.Textslyearopeningbalance.Enabled = True
                    master.Combosldiscountcategory.Enabled = True
                    master.Combosldistrictcode.Enabled = True
                    master.Textsladdress1.Enabled = True
                    master.Textsladdress2.Enabled = True
                    master.Textsladdress3.Enabled = True
                    master.Commandmastersave.Enabled = True
                    'master.Commandmasteredit.Enabled = False
                    'master.Commandmasteradd.Enabled = False
                    master.Commandmasterdelete.Enabled = True
                    master.Commandmasterabandon.Enabled = True
                    master.Textslsubledgerdiscription.SetFocus
                Else
                    master.sledger.Enabled = False
                
                End If
                
            End If
        End If
    End If
    
'///////////////********************////////
'       CREDIT NOTE END PART SEARCH
'///////////////********************////////
    If tablabel = 3 And Trim(UCase(masterlabel)) = UCase("master") Then
        
        master.Commandmasteradd.Enabled = True
        master.Commandmasteredit.Enabled = True
        
        Grid1.col = 0
        master.Combocnepcontragenledgerdesc.Text = Grid1.Text
        Grid1.col = 1
        master.Combocnepcontrasubledgerdesc.Text = Grid1.Text
        master.TextFINDSUBLEADGER = Grid1.Text
        Grid1.col = 2
        master.Combocnepgenledgerdesc.Text = Grid1.Text
        Grid1.col = 3
        master.Combocnepsubledgerdesc.Text = Grid1.Text
        Grid1.col = 4
        master.Textcnep20chartext.Text = Grid1.Text
        Grid1.col = 5
        master.Textcneprate.Text = Grid1.Text
        Grid1.col = 6
        master.Combocnepdrorcr.Text = Grid1.Text
        Grid1.col = 7
        master.CneTextInvePrintOrder.Text = Grid1.Text

        Unload Me
        master.Enabled = True
        If master.Commandmasteradd.Enabled = False Then
            master.Commandmastersave.Enabled = True
        End If
        If master.Commandmasteradd.Enabled = True Then
            If master.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                    master.CneTextInvePrintOrder.Enabled = False
                    master.Combocnepcontragenledgerdesc.Enabled = True
                    master.Combocnepcontrasubledgerdesc.Enabled = True
                    master.Combocnepgenledgerdesc.Enabled = True
                    master.Combocnepsubledgerdesc.Enabled = True
                    master.Textcnep20chartext.Enabled = True
                    master.Textcneprate.Enabled = True
                    master.Combocnepdrorcr.Enabled = True
                    master.Commandmastersave.Enabled = True
                    master.Commandmasteredit.Enabled = False
                    master.Commandmasteradd.Enabled = False
                    master.Commandmasterdelete.Enabled = True
                    master.Commandmasterabandon.Enabled = True
                    master.Combocnepcontragenledgerdesc.SetFocus
                End If
            End If
        End If
    End If

'**************** CASH END PART*******************

If tablabel = 5 And Trim(UCase(masterlabel)) = UCase("master") Then
        
        master.Commandmasteradd.Enabled = True
        master.Commandmasteredit.Enabled = True
        
        Grid1.col = 0
        master.cashCombocnepcontragenledgerdesc.Text = Grid1.Text
        Grid1.col = 1
        master.cashCombocnepcontrasubledgerdesc.Text = Grid1.Text
        master.TextFINDSUBLEADGER = Grid1.Text
        Grid1.col = 2
        master.cashCombocnepgenledgerdesc.Text = Grid1.Text
        Grid1.col = 3
        master.cashCombocnepsubledgerdesc.Text = Grid1.Text
        Grid1.col = 4
        master.cashTextcnep20chartext.Text = Grid1.Text
        Grid1.col = 5
        master.cashTextcneprate.Text = Grid1.Text
        Grid1.col = 6
        master.cashCombocnepdrorcr.Text = Grid1.Text
        Grid1.col = 7
        master.cashTextInvePrintOrder.Text = Grid1.Text

        Unload Me
        master.Enabled = True
        If master.Commandmasteradd.Enabled = False Then
            master.Commandmastersave.Enabled = True
        End If
        If master.Commandmasteradd.Enabled = True Then
            If master.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                    master.cashTextInvePrintOrder.Enabled = False
                    master.cashCombocnepcontragenledgerdesc.Enabled = True
                    master.cashCombocnepcontrasubledgerdesc.Enabled = True
                    master.cashCombocnepgenledgerdesc.Enabled = True
                    master.cashCombocnepsubledgerdesc.Enabled = True
                    master.cashTextcnep20chartext.Enabled = True
                    master.cashTextcneprate.Enabled = True
                    master.cashCombocnepdrorcr.Enabled = True
                    master.Commandmastersave.Enabled = True
                    master.Commandmasteredit.Enabled = False
                    master.Commandmasteradd.Enabled = False
                    master.Commandmasterdelete.Enabled = True
                    master.Commandmasterabandon.Enabled = True
                    master.cashCombocnepcontragenledgerdesc.SetFocus
                End If
            End If
        End If
    End If



'///////////////********************////////
'       INVOICE END PART SEARCH
'///////////////********************////////
    If tablabel = 2 And Trim(UCase(masterlabel)) = UCase("master") Then
         master.Commandmasteradd.Enabled = True
        master.Commandmasteredit.Enabled = True
        Grid1.col = 0
        master.Comboinvepcontragenledgerdesc.Text = Grid1.Text
        Grid1.col = 1
        master.Comboinvepcontrasubledgerdesc.Text = Grid1.Text
        master.TextFINDSUBLEADGER = Grid1.Text
        Grid1.col = 2
        master.Comboinvepgenledgerdesc.Text = Grid1.Text
        Grid1.col = 3
        master.Comboinvepsubledgerdesc.Text = Grid1.Text
        Grid1.col = 4
        master.Textinvep20chartext.Text = Grid1.Text
        Grid1.col = 5
        master.Textinveprate.Text = Grid1.Text
        Grid1.col = 6
        master.Comboinvepdrorcr.Text = Grid1.Text
        Grid1.col = 7
        master.TextInvePrintOrder.Text = Grid1.Text
        
        Unload Me
        master.Enabled = True
        If master.Commandmasteradd.Enabled = False Then
            master.Commandmastersave.Enabled = True
        End If
        If master.Commandmasteradd.Enabled = True Then
            If master.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                    master.TextInvePrintOrder.Enabled = False
                    master.invnoteend.Enabled = True
                    master.Comboinvepcontragenledgerdesc.Enabled = True
                    master.Comboinvepcontrasubledgerdesc.Enabled = True
                    master.Comboinvepgenledgerdesc.Enabled = True
                    master.Comboinvepsubledgerdesc.Enabled = True
                    master.Textinvep20chartext.Enabled = True
                    master.Textinveprate.Enabled = True
                    master.Comboinvepdrorcr.Enabled = True
                    master.Commandmastersave.Enabled = True
                    master.Commandmasteredit.Enabled = False
                    master.Commandmasteradd.Enabled = False
                    master.Commandmasterdelete.Enabled = True
                    master.Commandmasterabandon.Enabled = True
                    master.Comboinvepcontragenledgerdesc.SetFocus
                Else
                    
                    master.invnoteend.Enabled = False
                    
                End If
            End If
        End If
    End If

'************ FOR CASH /BANK

'If tablabel = 5 And Trim(UCase(masterlabel)) = UCase("master") Then
 '       master.Commandmasteradd.Enabled = True
  '      master.Commandmasteredit.Enabled = True
   '     Grid1.col = 0
      '  master.Textcbgenledger.Text = Grid1.Text
    '    Grid1.col = 1
     '   master.Textcbsubledger.Text = Grid1.Text
      '  master.TextFINDSUBLEADGER = Grid1.Text
       ' Unload Me
 '       master.Enabled = True
  '      If master.Commandmasteradd.Enabled = False Then
   '         master.Commandmastersave.Enabled = True
    '    End If
     '   If master.Commandmasteradd.Enabled = True Then
'            If master.Commandmasteredit.Enabled = True Then
 '               X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
  '              If X = 6 Then
   '                 master.CASHBANK.Enabled = True
    '                master.Textcbsubledger.Enabled = True
     '               master.Textcbgenledger.Enabled = True
      '              master.Commandmastersave.Enabled = True
       '             'master.Commandmasteredit.Enabled = False
        '            'master.Commandmasteradd.Enabled = False
         '           master.Commandmasterdelete.Enabled = True
          '          master.Commandmasterabandon.Enabled = True
           '         master.Textcbgenledger.SetFocus
            '    Else
             '       master.CASHBANK.Enabled = False
              '
               ' End If
                
  ''          End If
    '    End If
    'End If
'       BOOK SEARCH
'///////////////********************////////
          
        'by vk  bookmaster.Commandmasteradd.Enabled = True
                'bookmaster.Commandmasteredit.Enabled = True
          
    If tablabel = 0 And Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
        bookmaster.Commandmasteradd.Enabled = True
        bookmaster.Commandmasteredit.Enabled = True
        Grid1.col = 0
        bookmaster.Textbbookcode = Grid1.Text
        Grid1.col = 1
        bookmaster.Textbbookname.Text = Grid1.Text
        Grid1.col = 2
        bookmaster.Combobgroupcode.Text = Grid1.Text
        Grid1.col = 3
        bookmaster.Textbrate.Text = Grid1.Text
        Grid1.col = 4
        bookmaster.Textbdiscount.Text = Grid1.Text
        Unload Me
        bookmaster.Enabled = True
        bookmaster.Textfindbookcode.Text = bookmaster.Textbbookcode.Text
        If bookmaster.Commandmasteradd.Enabled = True Then
            If bookmaster.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                    bookmaster.booksmaster.Enabled = True
                    bookmaster.Textbbookcode.Enabled = False
                    'bookmaster.masteredit = True
                    bookmaster.Textfindbookcode.Text = bookmaster.Textbbookcode.Text
                    bookmaster.Textbbookname.Enabled = True
                    bookmaster.Combobgroupcode.Enabled = True
                    bookmaster.Combobgroupname.Enabled = True
                    bookmaster.Textbrate.Enabled = True
                    bookmaster.Textbdiscount.Enabled = True
                    bookmaster.Commandmastersave.Enabled = True
                    bookmaster.Commandmasterdelete.Enabled = True
                    bookmaster.Commandmasteredit.Enabled = True
                    bookmaster.Commandmasteradd.Enabled = True
                    bookmaster.Textbbookname.SetFocus
                Else
                
                  bookmaster.booksmaster.Enabled = False
                
                
                End If
            End If
        End If
    End If

'///////////////********************////////
'       BOOK GROUP SEARCH
'///////////////********************////////
    
    If tablabel = 1 And Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
        bookmaster.Commandmasteradd.Enabled = True
        bookmaster.Commandmasteredit.Enabled = True
        
        
        Grid1.col = 0
        bookmaster.Textbggroupcode.Text = Grid1.Text
        Grid1.col = 1
        bookmaster.Textbggroupname.Text = Grid1.Text
        Unload Me
        bookmaster.Enabled = True
        bookmaster.textbgfindcode.Text = bookmaster.Textbggroupcode.Text
        If bookmaster.Commandmasteradd.Enabled = True Then
            If bookmaster.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                    bookmaster.booksgroupmaster.Enabled = True
                    bookmaster.Textbggroupcode.Enabled = False
                    bookmaster.textbgfindcode.Text = bookmaster.Textbggroupcode.Text
                    bookmaster.Textbggroupname.Enabled = True
                    bookmaster.Commandmastersave.Enabled = True
                    bookmaster.Commandmasteredit.Enabled = False
                    bookmaster.Commandmasteradd.Enabled = False
                    bookmaster.Commandmasterdelete.Enabled = True
                    bookmaster.Textbggroupname.SetFocus
                    Else
                    bookmaster.booksgroupmaster.Enabled = False
                End If
            End If
        End If
    End If
'///////////////********************////////
'       Discount SEARCH
'///////////////********************////////
    
    If tablabel = 4 And Trim(UCase(masterlabel)) = Trim(UCase("master")) Then
        
        master.Commandmasteradd.Enabled = True
        master.Commandmasteredit.Enabled = True
        
        Grid1.col = 0
        master.Textdcdiscountcategorycode.Text = Grid1.Text
        master.Textfinddiscountcategory.Text = Grid1.Text
        Grid1.col = 1
        master.textfinddiscgroupcode.Text = Grid1.Text
        master.Combobgroupcode.Text = Grid1.Text
        Grid1.col = 2
        master.Textdcdiscountrate.Text = Grid1.Text
        Unload Me
        master.Enabled = True
        If master.Commandmasteradd.Enabled = False Then
            master.Commandmastersave.Enabled = True
        End If
        If master.Commandmasteradd.Enabled = True Then
            If master.Commandmasteredit.Enabled = True Then
               X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                    master.discount.Enabled = True
                    
                    master.Textdcdiscountcategorycode.Enabled = True
                    master.Combobgroupcode.Enabled = True
                    master.Combobgroupname.Enabled = True
                    master.Textdcdiscountrate.Enabled = True
                    master.Commandmasteredit.Enabled = True
                    master.Commandmasteradd.Enabled = True
                    master.Commandmasterabandon.Enabled = True
                    master.Commandmastersave.Enabled = True
                    master.Commandmasterdelete.Enabled = True
                    master.Textdcdiscountcategorycode.SetFocus
                    
                Else
                        master.discount.Enabled = False
                
                End If
                
            End If
        End If
    End If
'///////////////********************////////
'       Invoice  SEARCH
'///////////////********************////////
    
If tablabel = 11 And Trim(UCase(masterlabel)) = Trim(UCase("invoice")) Then
    INVOICE.invoiceabandon
    Grid1.col = 0
    INVOICE.I_NO = Grid1.Text
    INVOICE.Enabled = True
    INVOICE.edit = False
    Unload Me
    INVOICE.I_NO_LostFocus
    INVOICE.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
    Dim ctl As Control
    For Each ctl In INVOICE.Controls
        If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(INVOICE.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(INVOICE.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
        INVOICE.Commandadd.Enabled = True
        INVOICE.Commandedit.Enabled = True
        INVOICE.Commandsearch.Enabled = True
        INVOICE.Commandsave.Enabled = False
        INVOICE.Commanddelete.Enabled = True
        INVOICE.Commandabandon.Enabled = True
        INVOICE.CommandPrint.Enabled = True

    INVOICE.Picture5.Enabled = True
    addoredit = False
End If
'///////////////********************////////
'       CREDIT NOT ITEM  SEARCH
'///////////////********************////////
    
If tablabel = 13 And Trim(UCase(masterlabel)) = Trim(UCase("CREDITITEMNOTE")) Then
    CRITNOTE.CREDITAbandon
    Grid1.col = 0
    CRITNOTE.I_NO = Grid1.Text
    
    CRITNOTE.Enabled = True
    CRITNOTE.edit = False
    Unload Me
    CRITNOTE.I_NO_LostFocus
    CRITNOTE.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
    For Each ctl In CRITNOTE.Controls
        If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(CRITNOTE.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(CRITNOTE.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
        CRITNOTE.Commandadd.Enabled = True
        CRITNOTE.Commandedit.Enabled = True
        CRITNOTE.Commandsearch.Enabled = True
        CRITNOTE.Commandsave.Enabled = False
        CRITNOTE.Commanddelete.Enabled = True
        CRITNOTE.Commandabandon.Enabled = True
        CRITNOTE.CommandPrint.Enabled = True
        CRITNOTE.Picture5.Enabled = True
        CRITNOTE.customercode.Enabled = False
        addoredit = False
End If






'*******************************counter sale  search



If tablabel = 17 And Trim(UCase(masterlabel)) = Trim(UCase("countersale")) Then
    countersale.invoiceabandon
    Grid1.col = 0
     countersale.I_NO = Grid1.Text
     countersale.Enabled = True
     countersale.edit = False
    Unload Me
     countersale.I_NO_LostFocus
     countersale.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
    'Dim ctl As Control
    For Each ctl In countersale.Controls
        If Not TypeOf ctl Is CommandButton Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(countersale.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(INVOICE.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
         countersale.Commandadd.Enabled = True
         countersale.Commandedit.Enabled = True
         countersale.Commandsearch.Enabled = True
         countersale.Commandsave.Enabled = False
         countersale.Commanddelete.Enabled = True
         countersale.Commandabandon.Enabled = True
         countersale.CommandPrint.Enabled = True
         countersale.Picture5.Enabled = True
         addoredit = False
End If






'///////////////********************////////
'       voucher  SEARCH
'///////////////********************////////

    
    
    
If tablabel = 12 And Trim(UCase(masterlabel)) = Trim(UCase("voucher")) Then
   ' Voucherform.Commandabandon_Click
    'voucher.voucherabandon
    Grid1.col = 0
    Voucherform.vtype = Grid1.Text
    Grid1.col = 1
    Voucherform.vdate = Grid1.Text
    Grid1.col = 2
    Voucherform.vno = Grid1.Text
    Unload Me
            
    Voucherform.Enabled = True
    'Voucherform.edit = False
'    Voucherform.vtype.SetFocus
    'Voucherform.vtype_LostFocus
    Voucherform.vdate_LostFocus
    Voucherform.vno_LostFocus
    'Voucherform.vno.Enabled = False
    lastrow = 0
    lastcol = 1
'    For Each ctl In Voucherform.Controls
'        If Not TypeOf ctl Is CommandButton Then
 '               ctl.Enabled = False
  '      End If
'        If UCase(Trim(ctl.Name)) = UCase(Trim(Voucherform.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(Voucherform.Commandall.Name)) Then
'           ctl.Enabled = False
        'End If
   ' Next
  '  INVOICE.Picture5.Enabled = True
   ' addoredit = False
End If

'********************************
'  for Agent master search
'*******************************
If tablabel = 3 And Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
         bookmaster.Commandmasteradd.Enabled = True
        bookmaster.Commandmasteredit.Enabled = True
        
        Grid1.col = 0
        bookmaster.comboAgentMaster.Text = Grid1.Text
        
        Unload Me
        bookmaster.Enabled = True
        bookmaster.TextfindAgentmaster.Text = bookmaster.comboAgentMaster.Text
   
        If bookmaster.Commandmasteradd.Enabled = True Then
            If bookmaster.Commandmasteredit.Enabled = True Then
               ' X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
               ' If X = 6 Then
                        bookmaster.Agent.Enabled = True
                        bookmaster.comboAgentMaster.Enabled = False
                        bookmaster.TextfindAgentmaster.Text = bookmaster.comboAgentMaster.Text
                    
                        bookmaster.Addcd.Enabled = True
                        bookmaster.Removecd.Enabled = True
                    
                        bookmaster.ListDis1.Enabled = True
                        bookmaster.ListDis2.Enabled = True
                    
                        bookmaster.Commandmastersave.Enabled = True
                        bookmaster.Commandmasteredit.Enabled = True
                        bookmaster.Commandmasteradd.Enabled = True
                        bookmaster.Commandmasterdelete.Enabled = True
                        bookmaster.Agent.Enabled = True
                        bookmaster.comboAgentMaster.Enabled = True
                        bookmaster.comboAgentMaster.SetFocus
                'Else
                   
                 '       bookmaster.Agent.Enabled = False
                'End If
            End If
        End If
    End If
    
    
    '*************   for credit not search

 
If tablabel = 15 And Trim(UCase(masterlabel)) = Trim(UCase("Creditnotefile")) Then
        Grid1.col = 0
              
        Creditnotefile.Frame1.Enabled = True
       
        Creditnotefile.rs.Find "Cnn = " + Trim(Val(Grid1.Text)) + "", 1
         Unload Me
        If Not Creditnotefile.rs.EOF Then
          ' X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
             ' If X = 6 Then
                       Creditnotefile.Frame1.Enabled = True
                       'Creditnotefile.Commandedit_Click
                       
                       'Creditnotefile.Cbdate.SetFocus
                       Creditnotefile.Show
              'Else
                   '   Creditnotefile.Frame1.Enabled = False
                    '  Creditnotefile.Show
                
              'End If
         Else
          
                        Creditnotefile.Commandabandon_Click
                        Creditnotefile.Show
          
        End If

    End If



If tablabel = 16 And Trim(UCase(masterlabel)) = Trim(UCase("Debitnotefile")) Then
        Grid1.col = 0
              
        Debitnotefile.Frame1.Enabled = True
       
        Debitnotefile.rs.Find "dnn = " + Trim(Val(Grid1.Text)) + "", 1
         Unload Me
        If Not Debitnotefile.rs.EOF Then
           'X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
            '  If X = 6 Then
                       
                       Debitnotefile.Frame1.Enabled = True
                       'Debitnotefile.Commandedit_Click
                       Debitnotefile.Show
             ' Else
              '        Debitnotefile.Frame1.Enabled = False
               '       Debitnotefile.Show
                
             ' End If
           Else
          
                        'MsgBox "Record not found"
                        Debitnotefile.Commandabandon_Click
                        Debitnotefile.Show
          
          End If
          Debitnotefile.Show
    End If

End If


End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Grid1_DblClick
    Else
        X = KeyAscii
        textsearch.SetFocus
        SendKeys Chr(X)
    End If
End Sub
Function tempr(tb As Integer, master As String)
    tablabel = tb
    masterlabel = master
    'DATA1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
      
    If tb = 0 And Trim(UCase(master)) = Trim(UCase("master")) Then
         Grid1.ColWidth(0) = 1200
         Grid1.ColWidth(1) = 3000
         DATA1.RecordSource = "select * from GLEDGER order by gledger"
         DATA1.Refresh
    End If
    
    If tb = 1 And Trim(UCase(master)) = Trim(UCase("master")) Then
        ' Grid1.Rows = 6000
         Grid1.ColWidth(0) = 3000
         Grid1.ColWidth(1) = 3000
         rs.Open "select * from SLEDGER", con, adOpenStatic, adLockReadOnly, adCmdText
         Set Grid1.DataSource = rs
         Grid1.Refresh
    End If
    
'///////////////********************////////
'       CREDIT NOTE END PART SEARCH
'///////////////********************////////

    If tb = 3 And Trim(UCase(master)) = Trim(UCase("master")) Then
        DATA1.RecordSource = "select * from CREditEND order by CGENLEDGER"
        DATA1.Refresh
        Grid1.ColWidth(2) = 1600
        Grid1.ColWidth(4) = 1600

    End If
    
    
    
'///////////////********************////////
'       CASH  END PART SEARCH
'///////////////********************////////

    If tb = 5 And Trim(UCase(master)) = Trim(UCase("master")) Then
         DATA1.RecordSource = "select * from CASHEND order by CGENLEDGER"
         DATA1.Refresh
         Grid1.ColWidth(2) = 1600
         Grid1.ColWidth(4) = 1600

    End If
    
    
    
    
'///////////////********************////////
'       invoice END PART SEARCH
'///////////////********************////////

    If tb = 2 And Trim(UCase(master)) = Trim(UCase("master")) Then
         DATA1.RecordSource = "select * from invoiceend order by CGENLEDGER"
         DATA1.Refresh
         Grid1.ColWidth(2) = 1600
         Grid1.ColWidth(4) = 1600
    End If
    
    

'///////////////********************////////
'       BOOK SEARCH
'///////////////********************////////

    If tb = 0 And Trim(UCase(master)) = Trim(UCase("bookmaster")) Then
         Grid1.ColWidth(0) = 1000
         Grid1.ColWidth(1) = 3000
         DATA1.RecordSource = "select * from BOOKS order by bookcode"
         DATA1.Refresh
    End If
    
'///////////////********************////////
'       discount SEARCH
'///////////////********************////////
   
    If tb = 4 And Trim(UCase(master)) = Trim(UCase("master")) Then
        DATA1.RecordSource = "select * from disccats order by categorycode"
         DATA1.Refresh
        
     End If
    
     'If tb = 5 And Trim(UCase(master)) = Trim(UCase("master")) Then
      '   DATA1.RecordSource = "select * from cbmf order by gld,sld"
      '   Grid1.ColWidth(0) = 3500
       '  Grid1.ColWidth(1) = 4500
       '  DATA1.Refresh
    'End If
    
'///////////////********************////////
'       GROUP SEARCH
'///////////////********************////////
   
    If tb = 1 And Trim(UCase(master)) = Trim(UCase("bookmaster")) Then
         DATA1.RecordSource = "select * from Groups order by groupcode"
         DATA1.Refresh
    End If
    
'///////////////********************////////
'       invoice SEARCH
'///////////////********************////////
   
    If tb = 11 And Trim(UCase(master)) = Trim(UCase("invoice")) Then
        Grid1.ColWidth(2) = 4000
         
        DATA1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER,NETAMOUNT from invoicea order by invoiceno"
        DATA1.Refresh
        If DATA1.Recordset.RecordCount > 0 Then
          For I = 1 To DATA1.Recordset.RecordCount
           Grid1.row = I
           Grid1.col = 3
           Grid1.Text = Format(Grid1.Text, "0.00")
           Grid1.Refresh
         Next I
         
         Grid1.row = 1
         Grid1.col = 0
       End If
    End If
  
  
  '*******************************counter sale search
  
  
    If tb = 17 And Trim(UCase(master)) = Trim(UCase("countersale")) Then
        Grid1.ColWidth(2) = 4000
         
        DATA1.RecordSource = "select INVOICENO as CashMemoNo,INVOICEDATE as CashMDate  ,SUBLEDGER,NETAMOUNT as NetAmt from casha order by invoiceno"
        DATA1.Refresh
        If DATA1.Recordset.RecordCount > 0 Then
          For I = 1 To DATA1.Recordset.RecordCount
           Grid1.row = I
           Grid1.col = 3
           Grid1.Text = Format(Grid1.Text, "0.00")
           Grid1.Refresh
         Next I
         
         Grid1.row = 1
         Grid1.col = 0
       End If
    End If







'///////////////********************////////
'       voucher SEARCH
'///////////////********************////////
   
    If tb = 12 And Trim(UCase(master)) = Trim(UCase("voucher")) Then
        'Grid1.ColWidth(2) = 4000
        DATA1.RecordSource = "select * from vouchers where vouchertype = '" + Voucherform.vtype + "' order by voucherdate, vouchernumber"
        DATA1.Refresh
    End If
    
'///////////////********************////////
'       CREDIT ITEM NOTE
'///////////////********************////////
   
    If tb = 13 And Trim(UCase(master)) = Trim(UCase("CREDITITEMNOTE")) Then
        Grid1.ColWidth(2) = 4000
        DATA1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER,NETAMOUNT from CREDITA order by invoiceno"
        DATA1.Refresh
       If DATA1.Recordset.RecordCount > 0 Then
        For I = 1 To DATA1.Recordset.RecordCount
           Grid1.row = I
           Grid1.col = 3
           Grid1.Text = Format(Grid1.Text, "0.00")
           Grid1.Refresh
         Next I
         Grid1.row = 1
         Grid1.col = 0
       End If
 
         
         
     End If
    
    
    If tb = 3 And Trim(UCase(master)) = Trim(UCase("bookmaster")) Then
        Grid1.ColWidth(0) = 3000
        DATA1.RecordSource = "select distinct Agentname from agentmaster order by agentname"
        DATA1.Refresh
    End If
    '****** credit note file
    
    If tb = 15 And Trim(UCase(master)) = Trim(UCase("CREDITNOTEFILE")) Then
        Grid1.ColWidth(1) = 1000
        Grid1.ColWidth(2) = 5000
        Grid1.ColWidth(3) = 1000
        DATA1.RecordSource = "select Cnn as CreditNoteNo,cnd as CreditNoteDate,PSld as SubLedger ,na as NetAmount from Cnf1a order by Cnn,cnd"
        DATA1.Refresh
        If DATA1.Recordset.RecordCount > 0 Then
         For I = 1 To DATA1.Recordset.RecordCount
           Grid1.row = I
           Grid1.col = 3
           Grid1.Text = Format(Grid1.Text, "0.00")
           Grid1.Refresh
         Next I
         Grid1.row = 1
         Grid1.col = 0
         End If
            
    
    End If
    '************ debit note file
    If tb = 16 And Trim(UCase(master)) = Trim(UCase("DebitNOTEFILE")) Then
        Grid1.ColWidth(1) = 1000
        Grid1.ColWidth(2) = 5000
        Grid1.ColWidth(3) = 1000
        DATA1.RecordSource = "select dnn as DebitNoteNo,dnd as DebitNoteDate,PSld as SubLedger ,na as NetAmount   from dnfa order by dnn,dnd"
        DATA1.Refresh
        
        If DATA1.Recordset.RecordCount > 0 Then
         For I = 1 To DATA1.Recordset.RecordCount
           Grid1.row = I
           Grid1.col = 3
           Grid1.Text = Format(Grid1.Text, "0.00")
           Grid1.Refresh
         Next I
         Grid1.row = 1
         Grid1.col = 0
    End If
    
    End If
    
   '
    
    Me.Show
    'Grid1 = 300
    textsearch.SetFocus
End Function

Private Sub textsearch_Change()
    'Data1.Refresh
    If tablabel = 0 And Trim(UCase(masterlabel)) = UCase("master") Then
        Grid1.col = 1
        sets = True
    End If
    If tablabel = 1 And Trim(UCase(masterlabel)) = UCase("master") Then
        Grid1.col = 1
        sets = True
    End If
    If tablabel = 4 And Trim(UCase(masterlabel)) = UCase("master") Then
        Grid1.col = 0
        sets = True
    End If
    
    If tablabel = 5 And Trim(UCase(masterlabel)) = UCase("master") Then
        Grid1.col = 0
        sets = True
    End If
    
    
    
    'Data1.Recordset.FindFirst Trim(Data1.Recordset.Fields(grid1.col).Name) + " like '" + Trim(textsearch.Text) + "*'"
    'grid1.row = Data1.Recordset.AbsolutePosition + 1
    'grid1.SetFocus
    'SendKeys ("{LEFT}")
End Sub

Private Sub textsearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DATA1.Refresh
        If tablabel = 0 And Trim(UCase(masterlabel)) = UCase("master") Then
            Grid1.col = 1
            sets = True
        End If
        If tablabel = 1 And Trim(UCase(masterlabel)) = UCase("master") Then
            Grid1.col = 1
            sets = True
        End If
        If tablabel = 4 And Trim(UCase(masterlabel)) = UCase("master") Then
            Grid1.col = 0
            sets = True
        End If
        If tablabel = 11 And Trim(UCase(masterlabel)) = UCase("invoice") Then
             textsearch.Text = Val(textsearch.Text)
            Grid1.col = 0
            sets = True
        End If
        
        If tablabel = 12 And Trim(UCase(masterlabel)) = UCase("vouchers") Then
            Grid1.col = 0
            sets = True
        End If
        
        
       If tablabel = 15 And Trim(UCase(masterlabel)) = UCase("CreditnoteFile") Then
            
            Grid1.col = 0
            sets = True
      End If
      
      
      If tablabel = 17 And Trim(UCase(masterlabel)) = UCase("countersale") Then
            Grid1.col = 0
            sets = True
      End If

      

       If textsearch.Text <> "" Then
           If DATA1.Recordset.RecordCount > 0 Then
              If IsNumeric(Trim(textsearch.Text)) Then
                  DATA1.Recordset.Find Trim(DATA1.Recordset.Fields(Grid1.col).Name) + " like " + Trim(textsearch.Text) + ""
              Else
                   DATA1.Recordset.Find Trim(DATA1.Recordset.Fields(Grid1.col).Name) + " like '" + Trim(textsearch.Text) + "*'"
              End If
              If DATA1.Recordset.AbsolutePosition > 0 Then
                   Grid1.row = DATA1.Recordset.AbsolutePosition
              End If
           End If
             Grid1.SetFocus
               SendKeys ("{LEFT}")
           End If
    End If
End Sub
