VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form searchscreen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4650
   ClientLeft      =   2040
   ClientTop       =   2040
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Search.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Bindings        =   "Search.frx":000C
      Height          =   4215
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSMask.MaskEdBox textsearch 
      Height          =   285
      Left            =   1890
      TabIndex        =   0
      Top             =   4290
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3330
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "Search String"
      Height          =   225
      Left            =   330
      TabIndex        =   1
      Top             =   4320
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
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    GRID1.row = 0
    GRID1.col = 0
    Set con = New ADODB.Connection
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
    con.Open
    Set rs = New ADODB.Recordset
    'grid1.Rows = 1 'adarsh
    Me.Left = 1300
    Me.Top = 700
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
            End If
        End If
    End If
End Sub
Private Sub Grid1_DblClick()
    If Trim(UCase(masterlabel)) = UCase("master") Then
        For I = 0 To 5
            If master.SSTab1.Tab <> I Then
                master.SSTab1.TabEnabled(I) = False
            End If
        Next
    Else
        If Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
            For I = 0 To 2
                If bookmaster.SSTab1.Tab <> I Then
                    bookmaster.SSTab1.TabEnabled(I) = False
                End If
            Next
        End If
    End If
    
    
         master.Commandmasteradd.Enabled = True
         master.Commandmasteredit.Enabled = True
    If tablabel = 0 And Trim(UCase(masterlabel)) = UCase("master") Then
        GRID1.col = 0
        master.ComboSPECIALCATEGORY = GRID1.Text
        GRID1.col = 1
        master.Textglgeneralledgerdiscription = GRID1.Text
        master.Textfindgl.Text = GRID1.Text
        GRID1.col = 2
        If Trim(UCase(GRID1.Text)) = UCase("-1") Then
            master.GMASTERPL.Value = 1
        Else
            master.GMASTERPL.Value = 0
        End If
        GRID1.col = 3
        If Trim(UCase(GRID1.Text)) = UCase("-1") Then
            master.GMASTERBS.Value = 1
        Else
            master.GMASTERBS.Value = 0
        End If
        GRID1.col = 4
        If Trim(UCase(GRID1.Text)) = UCase("-1") Then
            master.GMASTERSL.Value = 1
        Else
            master.GMASTERSL.Value = 0
        End If
        GRID1.col = 5
        master.Textglyearopeningbalance = Trim(GRID1.Text)
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
        GRID1.col = 0
        master.Comboslgenledgerdiscription = GRID1.Text
        GRID1.col = 1
        master.Textslsubledgerdiscription = GRID1.Text
     
        master.TextFINDSUBLEADGER = GRID1.Text
        GRID1.col = 2
        master.Textsldiscriptionforinvoice = GRID1.Text
        GRID1.col = 3
        master.Textslyearopeningbalance = GRID1.Text
        GRID1.col = 4
        master.Combosldiscountcategory = GRID1.Text
        GRID1.col = 5
        master.Combosldistrictcode = GRID1.Text
        GRID1.col = 6
        master.Textsladdress1 = GRID1.Text
        GRID1.col = 7
        master.Textsladdress2 = GRID1.Text
        GRID1.col = 8
        master.Textsladdress3 = GRID1.Text
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
        GRID1.col = 0
        master.Combocnepcontragenledgerdesc.Text = GRID1.Text
        GRID1.col = 1
        master.Combocnepcontrasubledgerdesc.Text = GRID1.Text
        master.TextFINDSUBLEADGER = GRID1.Text
        GRID1.col = 2
        master.Combocnepgenledgerdesc.Text = GRID1.Text
        GRID1.col = 3
        master.Combocnepsubledgerdesc.Text = GRID1.Text
        GRID1.col = 4
        master.Textcnep20chartext.Text = GRID1.Text
        GRID1.col = 5
        master.Textcneprate.Text = GRID1.Text
        GRID1.col = 6
        master.Combocnepdrorcr.Text = GRID1.Text
        Unload Me
        master.Enabled = True
        If master.Commandmasteradd.Enabled = False Then
            master.Commandmastersave.Enabled = True
        End If
        If master.Commandmasteradd.Enabled = True Then
            If master.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
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
                End If
            End If
        End If
    End If
'///////////////********************////////
'       INVOICE END PART SEARCH
'///////////////********************////////
    If tablabel = 2 And Trim(UCase(masterlabel)) = UCase("master") Then
        GRID1.col = 0
        master.Comboinvepcontragenledgerdesc.Text = GRID1.Text
        GRID1.col = 1
        master.Comboinvepcontrasubledgerdesc.Text = GRID1.Text
        master.TextFINDSUBLEADGER = GRID1.Text
        GRID1.col = 2
        master.Comboinvepgenledgerdesc.Text = GRID1.Text
        GRID1.col = 3
        master.Comboinvepsubledgerdesc.Text = GRID1.Text
        GRID1.col = 4
        master.Textinvep20chartext.Text = GRID1.Text
        GRID1.col = 5
        master.Textinveprate.Text = GRID1.Text
        GRID1.col = 6
        master.Comboinvepdrorcr.Text = GRID1.Text
        Unload Me
        master.Enabled = True
        If master.Commandmasteradd.Enabled = False Then
            master.Commandmastersave.Enabled = True
        End If
        If master.Commandmasteradd.Enabled = True Then
            If master.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
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



'       BOOK SEARCH
'///////////////********************////////
          bookmaster.Commandmasteradd.Enabled = True
          bookmaster.Commandmasteredit.Enabled = True
          
    If tablabel = 0 And Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
        GRID1.col = 0
        bookmaster.Textbbookcode = GRID1.Text
        GRID1.col = 1
        bookmaster.Textbbookname.Text = GRID1.Text
        GRID1.col = 2
        bookmaster.Combobgroupcode.Text = GRID1.Text
        GRID1.col = 3
        bookmaster.Textbrate.Text = GRID1.Text
        GRID1.col = 4
        bookmaster.Textbdiscount.Text = GRID1.Text
        Unload Me
        bookmaster.Enabled = True
        bookmaster.Textfindbookcode.Text = bookmaster.Textbbookcode.Text
        If bookmaster.Commandmasteradd.Enabled = True Then
            If bookmaster.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                    bookmaster.booksmaster.Enabled = True
                    bookmaster.Textbbookcode.Enabled = False
                    
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
        GRID1.col = 0
        bookmaster.Textbggroupcode.Text = GRID1.Text
        GRID1.col = 1
        bookmaster.Textbggroupname.Text = GRID1.Text
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
        GRID1.col = 0
        master.Textdcdiscountcategorycode.Text = GRID1.Text
        master.Textfinddiscountcategory.Text = GRID1.Text
        GRID1.col = 1
        master.textfinddiscgroupcode.Text = GRID1.Text
        master.Combobgroupcode.Text = GRID1.Text
        GRID1.col = 2
        master.Textdcdiscountrate.Text = GRID1.Text
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
    GRID1.col = 0
    INVOICE.I_NO = GRID1.Text
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
    GRID1.col = 0
    CRITNOTE.I_NO = GRID1.Text
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
        addoredit = False
End If



'///////////////********************////////
'       voucher  SEARCH
'///////////////********************////////


If tablabel = 12 And Trim(UCase(masterlabel)) = Trim(UCase("voucher")) Then
   ' Voucherform.Commandabandon_Click
    'voucher.voucherabandon
    GRID1.col = 0
    Voucherform.vtype = GRID1.Text
    GRID1.col = 1
    Voucherform.vdate = GRID1.Text
    GRID1.col = 2
    Voucherform.vno = GRID1.Text
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
 '   DATA1.DatabaseName = VB.App.Path + "\" + main.directory + "\DATA.MDB"
          
    If tb = 0 And Trim(UCase(master)) = Trim(UCase("master")) Then
        DATA1.RecordSource = "select * from GLEDGER order by gledger"
    End If
    
    If tb = 1 And Trim(UCase(master)) = Trim(UCase("master")) Then
        DATA1.RecordSource = "select * from SLEDGER order by SUBLEDGER"
    End If
    
'///////////////********************////////
'       CREDIT NOTE END PART SEARCH
'///////////////********************////////

    If tb = 3 And Trim(UCase(master)) = Trim(UCase("master")) Then
        DATA1.RecordSource = "select * from CRENOTEEND order by CGENLEDGER"
    End If
    
'///////////////********************////////
'       invoice END PART SEARCH
'///////////////********************////////

    If tb = 2 And Trim(UCase(master)) = Trim(UCase("master")) Then
        DATA1.RecordSource = "select * from invoiceend order by CGENLEDGER"
    End If
    
'///////////////********************////////
'       BOOK SEARCH
'///////////////********************////////

    If tb = 0 And Trim(UCase(master)) = Trim(UCase("bookmaster")) Then
        DATA1.RecordSource = "select * from BOOKS order by bookcode"
    End If
    
'///////////////********************////////
'       discount SEARCH
'///////////////********************////////
   
    If tb = 4 And Trim(UCase(master)) = Trim(UCase("master")) Then
        DATA1.RecordSource = "select * from disccats order by categorycode"
    End If
    
     If tb = 5 And Trim(UCase(master)) = Trim(UCase("master")) Then
        DATA1.RecordSource = "select * from disccats order by categorycode"
    End If
    
'///////////////********************////////
'       GROUP SEARCH
'///////////////********************////////
   
    If tb = 1 And Trim(UCase(master)) = Trim(UCase("bookmaster")) Then
        DATA1.RecordSource = "select * from Groups order by groupcode"
    End If
'///////////////********************////////
'       invoice SEARCH
'///////////////********************////////
   
    If tb = 11 And Trim(UCase(master)) = Trim(UCase("invoice")) Then
        GRID1.ColWidth(2) = 4000
        DATA1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER,NETAMOUNT from invoicea order by invoiceno"
    End If
'///////////////********************////////
'       voucher SEARCH
'///////////////********************////////
   
    If tb = 12 And Trim(UCase(master)) = Trim(UCase("voucher")) Then
        'Grid1.ColWidth(2) = 4000
        DATA1.RecordSource = "select * from vouchers where vouchertype = '" + Voucherform.vtype + "' order by voucherdate, vouchernumber"
    End If
    
'///////////////********************////////
'       CREDIT ITEM NOTE
'///////////////********************////////
   
    If tb = 13 And Trim(UCase(master)) = Trim(UCase("CREDITITEMNOTE")) Then
        GRID1.ColWidth(2) = 4000
        DATA1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER,NETAMOUNT from CREDITA order by invoiceno"
    End If
    
  '  DATA1.Refresh
    Me.Show
    textsearch.SetFocus
End Function
Private Sub textsearch_Change()
    'Data1.Refresh
    If tablabel = 0 And Trim(UCase(masterlabel)) = UCase("master") Then
        GRID1.col = 1
        sets = True
    End If
    If tablabel = 1 And Trim(UCase(masterlabel)) = UCase("master") Then
        GRID1.col = 1
        sets = True
    End If
    If tablabel = 4 And Trim(UCase(masterlabel)) = UCase("master") Then
        GRID1.col = 0
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
            GRID1.col = 1
            sets = True
        End If
        If tablabel = 1 And Trim(UCase(masterlabel)) = UCase("master") Then
            GRID1.col = 1
            sets = True
        End If
        If tablabel = 4 And Trim(UCase(masterlabel)) = UCase("master") Then
            GRID1.col = 0
            sets = True
        End If
        If tablabel = 11 And Trim(UCase(masterlabel)) = UCase("invoice") Then
            GRID1.col = 0
            sets = True
        End If
        If tablabel = 12 And Trim(UCase(masterlabel)) = UCase("vouchers") Then
            GRID1.col = 0
            sets = True
        End If
        If DATA1.Recordset.RecordCount > 0 Then
           DATA1.Recordset.FindFirst Trim(DATA1.Recordset.Fields(GRID1.col).Name) + " like '" + Trim(textsearch.Text) + "*'"
           GRID1.row = DATA1.Recordset.AbsolutePosition + 1
           GRID1.SetFocus
           SendKeys ("{LEFT}")
        End If
    End If
End Sub
