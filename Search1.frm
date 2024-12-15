VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form searchscreen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4605
   ClientLeft      =   2040
   ClientTop       =   2040
   ClientWidth     =   8940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Search1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid Grid1 
      Bindings        =   "Search1.frx":000C
      Height          =   3975
      Left            =   60
      Negotiate       =   -1  'True
      TabIndex        =   2
      Top             =   60
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
            Type            =   1
            Format          =   """$""#,##0.00"
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   4560
      Top             =   2520
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox textsearch 
      Height          =   285
      Left            =   1530
      TabIndex        =   0
      Top             =   4170
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
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   1185
   End
End
Attribute VB_Name = "searchscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim RS As ADODB.Recordset
Dim tablabel As Integer
Dim masterlabel As String
Dim maxrow As Integer
Dim ctl As Control
Dim intselcol As Integer
Dim blnsortacr As Boolean
Dim strsortorder As String


Private Sub Agent1_ActivateInput(ByVal CharacterID As String)

End Sub

Private Sub Combo1_Change()

End Sub

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
                    invoice.Enabled = True
                End If
                If Trim(UCase(masterlabel)) = Trim(UCase("purchase")) Then
                    frmPurchase.Enabled = True
                End If
                
                If Trim(UCase(masterlabel)) = Trim(UCase("credititemnote")) Then
                    crtitem.Enabled = True
                End If
                
                If Trim(UCase(masterlabel)) = Trim(UCase("countersale")) Then
                    countersale.Enabled = True
                End If
                
               If Trim(UCase(masterlabel)) = Trim(UCase("Debitnotefile")) Then

                   Debitnotefile.Enabled = True
                
              End If
              If Trim(UCase(masterlabel)) = Trim(UCase("Creditnotefile")) Then
                    
                    Creditnotefile.Enabled = True
              End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Set RS = New ADODB.Recordset
    Data1.ConnectionString = CON
    Me.Left = 145
    Me.Top = 600
    intselcol = 0
    blnsortacr = True
    strsortorder = " asc"
    End Sub

Private Sub Form_Terminate()
    If Trim(UCase(masterlabel)) = Trim(UCase("master")) Then
        master.Enabled = True
    Else
        If Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
            bookmaster.Enabled = True
        Else
            If Trim(UCase(masterlabel)) = Trim(UCase("invoice")) Then
                invoice.Enabled = True
            End If
            
            If Trim(UCase(masterlabel)) = Trim(UCase("countersale")) Then
                countersale.Enabled = True
            End If

            
        End If
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    If Trim(UCase(masterlabel)) = Trim(UCase("master")) Then
        master.Enabled = True
    Else
        If Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
            bookmaster.Enabled = True
        Else
            If Trim(UCase(masterlabel)) = Trim(UCase("invoice")) Then
                invoice.Enabled = True
            Else
               If Trim(UCase(masterlabel)) = Trim(UCase("voucher")) Then
                    Voucherform.Enabled = True
               End If
                If Trim(UCase(masterlabel)) = Trim(UCase("CREDITITEMNOTE")) Then
                    crtitem.Enabled = True
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

Private Sub Grid1_Click()
'MsgBox DATA1.Recordset.Fields(Grid1.SelEndCol).Type
'Exit Sub
If Grid1.SelEndCol = intselcol And blnsortacr = True Then
strsortorder = " desc"
blnsortacr = False
ElseIf Grid1.SelEndCol = intselcol And blnsortacr = False Then
strsortorder = " asc"
blnsortacr = True
End If
intselcol = Grid1.SelEndCol
If intselcol >= 0 Then
Data1.Recordset.Sort = Data1.Recordset.Fields(intselcol).Name & strsortorder
End If
'MsgBox DATA1.Recordset.Fields(intselcol).Type
End Sub

Private Sub Grid1_DblClick()
'On Error GoTo 999

On Error Resume Next

If Data1.Recordset.RecordCount > 0 Then
    If Trim(UCase(masterlabel)) = UCase("master") Then
        For I = 0 To 5
            If master.SStab1.Tab <> I Then
                master.SStab1.TabEnabled(I) = False
            End If
        Next
    Else
        If Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
            For I = 0 To 3
                If bookmaster.SStab1.Tab <> I Then
                    bookmaster.SStab1.TabEnabled(I) = False
                End If
            Next
        End If
    End If
    
    
'by vk  'master.Commandmasteradd.Enabled = True
        'master.Commandmasteredit.Enabled = True
    If tablabel = 0 And Trim(UCase(masterlabel)) = UCase("master") Then
        
        master.Commandmasteradd.Enabled = True
        master.Commandmasteredit.Enabled = True
        
        Grid1.Col = 0
        master.ComboSPECIALCATEGORY = Grid1.Text
        Grid1.Col = 1
        master.Textglgeneralledgerdiscription = Grid1.Text
        master.Textfindgl.Text = Grid1.Text
        Grid1.Col = 2
        If Grid1.Text = True Then
            master.GMASTERPL.value = 1
        Else
            master.GMASTERPL.value = 0
        End If
        Grid1.Col = 3
        If Grid1.Text = True Then
            master.GMASTERBS.value = 1
        Else
            master.GMASTERBS.value = 0
        End If
        
        Grid1.Col = 4
        If Grid1.Text = True Then
            master.GMASTERSL.value = 1
        Else
            master.GMASTERSL.value = 0
        End If
        Grid1.Col = 5
        master.Textglyearopeningbalance = Format(Trim(Grid1.Text), "0.00")
        Grid1.Col = 6
        If Grid1.Text = True Then
            master.Cashbankbook.value = 1
        Else
            master.Cashbankbook.value = 0
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
                    master.Commandmasteredit_Click
''                    master.Textglyearopeningbalance.Enabled = True
''                    master.GMASTERPL.Enabled = True
''                    master.GMASTERBS.Enabled = True
''                    master.GMASTERSL.Enabled = True
''                    master.Cashbankbook.Enabled = True
''                    master.Textglgeneralledgerdiscription.Enabled = False
''                    master.ComboSPECIALCATEGORY.Enabled = True
''                    master.Commandmastersave.Enabled = True
''                    master.Commandmasteredit.Enabled = True
''                    master.Commandmasteradd.Enabled = False
''                    master.Commandmasterdelete.Enabled = False
''                    master.Commandmasterabandon.Enabled = True
''                  ' master.Textglgeneralledgerdiscription.Enabled = True
''                    master.Textglyearopeningbalance.Enabled = True
''                    master.ComboSPECIALCATEGORY.SetFocus
''                    'SetButton master.Commandmasteradd, master.Commandmasteredit, master.Commandmastersave, master.Commandmasterdelete
                Else
                    master.gledger.Enabled = False
                    master.Commandmastersave.Enabled = False
                    master.Commandmasteredit.Enabled = True
                    master.Commandmasteradd.Enabled = True
                    master.Commandmasterdelete.Enabled = True
                    master.Commandmasterabandon.Enabled = True
                  
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
        Grid1.Col = 0
        master.Comboslgenledgerdiscription = Grid1.Text
        Grid1.Col = 5
        If Grid1.Text <> "" Then
        Grid1.Col = 1
        master.Textslsubledgerdiscription = Right(Grid1.Text, Len(Grid1.Text) - InStr(1, Grid1.Text, " "))
        master.txtdistcode.Caption = Trim(Left(Grid1.Text, InStr(1, Grid1.Text, " ")))
        Else
        Grid1.Col = 1
        master.Textslsubledgerdiscription = Grid1.Text
        End If
        'master.CBODISTCODE.Text = ""
        'If InStr(1, Grid1.Text, "-") <> 0 Then
        'master.Combosldistrictcode.Text = Left(Grid1.Text, InStr(1, Grid1.Text, "-") - 1)
        'Else
        'master.Combosldistrictcode.ListIndex = 0
        'End If
        'If InStr(1, Grid1.Text, "-") <> 0 Then
        'master.TXTCUSTCODE.Caption = Mid(Grid1.Text, InStr(1, Grid1.Text, "-") + 1, 3)
        'Else
        'master.TXTCUSTCODE.Caption = ""
        'End If
        master.TextFINDSUBLEADGER = Grid1.Text
        Grid1.Col = 2
        master.Textsldiscriptionforinvoice = Grid1.Text
        Grid1.Col = 3
        master.Textslyearopeningbalance = Format(Trim(Grid1.Text), "0.00")
        Grid1.Col = 4
        master.Combosldiscountcategory = Grid1.Text
        Grid1.Col = 5
        If Grid1.Text <> "" Then
        master.Combosldistrictcode = Grid1.Text
        Else
        master.Combosldistrictcode.ListIndex = 0
        End If
        Grid1.Col = 6
        master.Textsladdress1 = Grid1.Text
        Grid1.Col = 7
        master.Textsladdress2 = Grid1.Text
        Grid1.Col = 8
        master.Textsladdress3 = Grid1.Text
        Grid1.Col = 9
        master.txtphoneno.Text = Grid1.Text
        
        Unload Me
        master.Enabled = True
        If master.Commandmasteradd.Enabled = False Then
            master.Commandmastersave.Enabled = True
        End If
        If master.Commandmasteradd.Enabled = True Then
            If master.Commandmasteredit.Enabled = True Then
                X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                master.addmaster = False
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
                    master.txtphoneno.Enabled = True
                    master.Commandmastersave.Enabled = True
                    master.Commandmasteredit.Enabled = False
                    master.Commandmasteradd.Enabled = False
                    master.Commandmasterdelete.Enabled = False
                    master.Commandmasterabandon.Enabled = True
                    master.Textslsubledgerdiscription.SetFocus
                    'SetButton master.Commandmasteradd, master.Commandmasteredit, master.Commandmastersave, master.Commandmasterdelete
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
        
        Grid1.Col = 0
        master.Combocnepcontragenledgerdesc.Text = Grid1.Text
        Grid1.Col = 1
        master.Combocnepcontrasubledgerdesc.Text = Grid1.Text
        master.TextFINDSUBLEADGER = Grid1.Text
        Grid1.Col = 2
        master.Combocnepgenledgerdesc.Text = Grid1.Text
        Grid1.Col = 3
        master.Combocnepsubledgerdesc.Text = Grid1.Text
        Grid1.Col = 4
        master.Textcnep20chartext.Text = Grid1.Text
        Grid1.Col = 5
        master.Textcneprate.Text = Format(Grid1.Text, "0.00")
        
        Grid1.Col = 6
        master.Combocnepdrorcr.Text = Grid1.Text
        Grid1.Col = 7
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
                    SetButton master.Commandmasteradd, master.Commandmasteredit, master.Commandmastersave, master.Commandmasterdelete
                    master.Combocnepcontragenledgerdesc.SetFocus
                    
                End If
            End If
        End If
    End If

'**************** CASH END PART*******************

If tablabel = 5 And Trim(UCase(masterlabel)) = UCase("master") Then
        
        master.Commandmasteradd.Enabled = True
        master.Commandmasteredit.Enabled = True
        
        Grid1.Col = 0
        master.cashCombocnepcontragenledgerdesc.Text = Grid1.Text
        Grid1.Col = 1
        master.cashCombocnepcontrasubledgerdesc.Text = Grid1.Text
        master.TextFINDSUBLEADGER = Grid1.Text
        Grid1.Col = 2
        master.cashCombocnepgenledgerdesc.Text = Grid1.Text
        Grid1.Col = 3
        master.cashCombocnepsubledgerdesc.Text = Grid1.Text
        Grid1.Col = 4
        master.cashTextcnep20chartext.Text = Grid1.Text
        Grid1.Col = 5
        master.cashTextcneprate.Text = Format(Grid1.Text, "0.00")
        Grid1.Col = 6
        master.cashCombocnepdrorcr.Text = Grid1.Text
        Grid1.Col = 7
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
                    master.cashend.Enabled = True
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
                    SetButton master.Commandmasteradd, master.Commandmasteredit, master.Commandmastersave, master.Commandmasterdelete
                Else
                    master.cashend.Enabled = False
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
        Grid1.Col = 0
        master.Comboinvepcontragenledgerdesc.Text = Grid1.Text
        Grid1.Col = 1
        master.Comboinvepcontrasubledgerdesc.Text = Grid1.Text
        master.TextFINDSUBLEADGER = Grid1.Text
        Grid1.Col = 2
        master.Comboinvepgenledgerdesc.Text = Grid1.Text
        Grid1.Col = 3
        master.Comboinvepsubledgerdesc.Text = Grid1.Text
        Grid1.Col = 4
        master.Textinvep20chartext.Text = Grid1.Text
        Grid1.Col = 5
        master.Textinveprate.Text = Format(Grid1.Text, "0.00")
        Grid1.Col = 6
        master.Comboinvepdrorcr.Text = Grid1.Text
        Grid1.Col = 7
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
                    SetButton master.Commandmasteradd, master.Commandmasteredit, master.Commandmastersave, master.Commandmasterdelete
                    master.Comboinvepcontragenledgerdesc.SetFocus
                    
                Else
                    
                    master.invnoteend.Enabled = False
                    
                End If
            End If
        End If
    End If

'************ FOR CASH /BANK

          
    If tablabel = 0 And Trim(UCase(masterlabel)) = Trim(UCase("bookmaster")) Then
        bookmaster.Commandmasteradd.Enabled = True
        bookmaster.Commandmasteredit.Enabled = True
        Grid1.Col = 0
        bookmaster.Combobgroupcode.Text = Left(Grid1.Text, 2)
        bookmaster.Textbbookcode = Right(Grid1.Text, Len(Grid1.Text) - 2)
        Grid1.Col = 1
        bookmaster.Textbbookname.Text = Grid1.Text
        Grid1.Col = 2
        bookmaster.txtSize1 = Grid1.Text
        Grid1.Col = 3
        bookmaster.txtunit1 = Grid1.Text
        Grid1.Col = 4
        bookmaster.txtSize2 = Grid1.Text
        Grid1.Col = 5
        bookmaster.txtunit2 = Grid1.Text
        Grid1.Col = 6
        bookmaster.txtQuality = Grid1.Text
        Grid1.Col = 7
        bookmaster.txtper = Grid1.Text
        Grid1.Col = 8
        bookmaster.Textbrate.Text = Grid1.Text
        Grid1.Col = 9
        bookmaster.cboready.Text = IIf(Grid1.Text = "0" Or Grid1.Text = "", "No", "Yes") 'rs.Fields(0).Value
        
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
                    bookmaster.txtSize1.Enabled = True
                    bookmaster.txtunit1.Enabled = True
                    bookmaster.txtSize2.Enabled = True
                    bookmaster.txtunit2.Enabled = True
                    bookmaster.txtQuality.Enabled = True
                    'bookmaster.Combobgroupcode.Enabled = True
                    'bookmaster.Combobgroupname.Enabled = True
                    bookmaster.Textbrate.Enabled = True
                    bookmaster.txtper.Enabled = True
                    'bookmaster.Textbdiscount.Enabled = True
                    bookmaster.Commandmastersave.Enabled = True
                    bookmaster.Commandmasterdelete.Enabled = True
                    bookmaster.Commandmasteredit.Enabled = True
                    bookmaster.Commandmasteradd.Enabled = True
                    SetButton bookmaster.Commandmasteradd, bookmaster.Commandmasteredit, bookmaster.Commandmastersave, bookmaster.Commandmasterdelete
                    bookmaster.Textbbookname.SetFocus
                    
                    'For Each ctl In bookmaster.Controls
                    'If UCase(ctl.Container.Name) = UCase("bookmaster") Then
                     '   If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Or TypeOf ctl Is ListBox Then
                      '  ctl.Enabled = True
                       ' End If
                    'End If
                    'Next
                    
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
        
        
        Grid1.Col = 0
        bookmaster.Textbggroupcode.Text = Grid1.Text
        Grid1.Col = 1
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
                    SetButton bookmaster.Commandmasteradd, bookmaster.Commandmasteredit, bookmaster.Commandmastersave, bookmaster.Commandmasterdelete
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
        
        Grid1.Col = 0
        master.Textdcdiscountcategorycode.Text = Grid1.Text
        master.Textfinddiscountcategory.Text = Grid1.Text
        Grid1.Col = 1
        master.textfinddiscgroupcode.Text = Grid1.Text
        master.Combobgroupcode.Text = Grid1.Text
        Grid1.Col = 2
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
                    master.DISCOUNT.Enabled = True
                    master.Textdcdiscountcategorycode.Enabled = True
                    master.Combobgroupcode.Enabled = True
                    master.Combobgroupname.Enabled = True
                    master.Textdcdiscountrate.Enabled = True
                    master.Commandmasteredit.Enabled = True
                    master.Commandmasteradd.Enabled = True
                    master.Commandmasterabandon.Enabled = True
                    master.Commandmastersave.Enabled = True
                    master.Commandmasterdelete.Enabled = True
                    SetButton master.Commandmasteradd, master.Commandmasteredit, master.Commandmastersave, master.Commandmasterdelete
                    master.Textdcdiscountcategorycode.SetFocus
                     
                Else
                   master.DISCOUNT.Enabled = False
                
                End If
                
            End If
        End If
    End If
'///////////////********************////////
'       Invoice  SEARCH
'///////////////********************////////
If tablabel = 13 And Trim(UCase(masterlabel)) = Trim(UCase("Purchase")) Then
    frmPurchase.invoiceabandon
    Grid1.Col = 0
    frmPurchase.I_NO = Grid1.Text
    frmPurchase.Enabled = True
    frmPurchase.Edit = False
    Unload Me
    frmPurchase.I_NO_LostFocus
    frmPurchase.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
'    Dim ctl As Control
    For Each ctl In frmPurchase.Controls
        If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is Timer And Not TypeOf ctl Is CrystalReport Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(frmPurchase.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(frmPurchase.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    frmPurchase.Commandadd.Enabled = True
    frmPurchase.Commandedit.Enabled = True
    frmPurchase.Commandsearch.Enabled = True
    frmPurchase.Commandsave.Enabled = False
    frmPurchase.Commanddelete.Enabled = True
    frmPurchase.Commandabandon.Enabled = True
    frmPurchase.CommandPrint.Enabled = True
    frmPurchase.Picture5.Enabled = True
    addoredit = False
End If

If tablabel = 11 And Trim(UCase(masterlabel)) = Trim(UCase("invoice")) Then
    invoice.invoiceabandon
    Grid1.Col = 0
    invoice.I_NO = Grid1.Text
    invoice.Enabled = True
    invoice.Edit = False
    Unload Me
    invoice.I_NO_LostFocus
    invoice.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
'    Dim ctl As Control
    For Each ctl In invoice.Controls
        If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is Timer And Not TypeOf ctl Is CrystalReport Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(invoice.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(invoice.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    invoice.Commandadd.Enabled = True
    invoice.Commandedit.Enabled = True
    invoice.Commandsearch.Enabled = True
    invoice.Commandsave.Enabled = False
    invoice.Commanddelete.Enabled = True
    invoice.Commandabandon.Enabled = True
    invoice.CommandPrint.Enabled = True
    invoice.Picture5.Enabled = True
    addoredit = False
End If


If tablabel = 111 And Trim(UCase(masterlabel)) = Trim(UCase("challan")) Then
    InvoiceChallane.invoiceabandon
    Grid1.Col = 0
    InvoiceChallane.I_NO = Grid1.Text
    InvoiceChallane.Enabled = True
    InvoiceChallane.Edit = False
    Unload Me
    InvoiceChallane.I_NO_LostFocus
    InvoiceChallane.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
    For Each ctl In InvoiceChallane.Controls
        If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is Timer And Not TypeOf ctl Is CrystalReport Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(InvoiceChallane.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(InvoiceChallane.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    InvoiceChallane.Commandadd.Enabled = True
    InvoiceChallane.Commandedit.Enabled = True
    InvoiceChallane.Commandsearch.Enabled = True
    InvoiceChallane.Commandsave.Enabled = False
    InvoiceChallane.Commanddelete.Enabled = True
    InvoiceChallane.Commandabandon.Enabled = True
    InvoiceChallane.CommandPrint.Enabled = True
    InvoiceChallane.Picture5.Enabled = True
    addoredit = False
End If


If tablabel = 112 And Trim(UCase(masterlabel)) = Trim(UCase("challannew")) Then
    frmChallan.invoiceabandon
    Grid1.Col = 0
    frmChallan.I_NO = Grid1.Text
    frmChallan.Enabled = True
    frmChallan.Edit = False
    Unload Me
    frmChallan.I_NO_LostFocus
    frmChallan.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
    For Each ctl In frmChallan.Controls
        If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is Timer And Not TypeOf ctl Is CrystalReport Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(frmChallan.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(frmChallan.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    frmChallan.Commandadd.Enabled = True
    frmChallan.Commandedit.Enabled = True
    frmChallan.Commandsearch.Enabled = True
    frmChallan.Commandsave.Enabled = False
    frmChallan.Commanddelete.Enabled = True
    frmChallan.Commandabandon.Enabled = True
    frmChallan.CommandPrint.Enabled = True
    frmChallan.Picture5.Enabled = True
    addoredit = False
End If


If tablabel = 12 And Trim(UCase(masterlabel)) = Trim(UCase("IssueA")) Then
    frmIssueItem.invoiceabandon
    Grid1.Col = 0
    frmIssueItem.I_NO = Grid1.Text
    frmIssueItem.Enabled = True
    frmIssueItem.Edit = False
    Unload Me
    frmIssueItem.I_NO_LostFocus
    frmIssueItem.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
'    Dim ctl As Control
    For Each ctl In frmIssueItem.Controls
        If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is Timer And Not TypeOf ctl Is CrystalReport Then
                ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(frmIssueItem.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(frmIssueItem.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    frmIssueItem.Commandadd.Enabled = True
    frmIssueItem.Commandedit.Enabled = True
    frmIssueItem.Commandsearch.Enabled = True
    frmIssueItem.Commandsave.Enabled = False
    frmIssueItem.Commanddelete.Enabled = True
    frmIssueItem.Commandabandon.Enabled = True
    frmIssueItem.CommandPrint.Enabled = True
    frmIssueItem.Picture5.Enabled = True
    addoredit = False
End If

'///////////////********************////////
'       CREDIT NOT ITEM  SEARCH
'///////////////********************////////
    
If tablabel = 13 And Trim(UCase(masterlabel)) = Trim(UCase("CREDITITEMNOTE")) Then
    crtitem.invoiceabandon
    'crtitem.CREDITAbandon
    Grid1.Col = 0
    crtitem.I_NO = Grid1.Text
    
    crtitem.Enabled = True
    crtitem.Edit = False
    Unload Me
    crtitem.I_NO_LostFocus
    crtitem.I_NO.Enabled = False
    lastrow = 0
    lastcol = 1
    
    'Dim ctl As Control
    For Each ctl In crtitem.Controls
        If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is Timer And Not TypeOf ctl Is CrystalReport Then
             ctl.Enabled = False
        End If
        If UCase(Trim(ctl.Name)) = UCase(Trim(crtitem.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(crtitem.Commandall.Name)) Then
           ctl.Enabled = False
        End If
    Next
    
    'For Each ctl In crtitem.Controls
     '   If Not TypeOf ctl Is CommandButton Then
      '          ctl.Enabled = False
       ' End If
       ' If UCase(Trim(ctl.Name)) = UCase(Trim(crtitem.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(crtitem.Commandall.Name)) Then
       '    ctl.Enabled = False
       ' End If
    'Next
        crtitem.Picture5.Enabled = True
        crtitem.Commandadd.Enabled = True
        crtitem.Commandedit.Enabled = True
        crtitem.Commandsearch.Enabled = True
        crtitem.Commandsave.Enabled = False
        crtitem.Commanddelete.Enabled = True
        crtitem.Commandabandon.Enabled = True
        crtitem.CommandPrint.Enabled = True
        crtitem.Picture5.Enabled = True
        crtitem.customercode.Enabled = False
        'crtitem.Picture5.Enabled = True
        addoredit = False
End If



'*******************************counter sale  search



If tablabel = 17 And Trim(UCase(masterlabel)) = Trim(UCase("countersale")) Then
    countersale.invoiceabandon
    Grid1.Col = 0
    countersale.I_NO = Grid1.Text
    countersale.Enabled = True
    countersale.Edit = False
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
        If UCase(Trim(ctl.Name)) = UCase(Trim(countersale.Commandother.Name)) Or UCase(Trim(ctl.Name)) = UCase(Trim(countersale.Commandall.Name)) Then
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
    Grid1.Col = 0
    Voucherform.vtype = Grid1.Text
    Grid1.Col = 1
    Voucherform.vdate = Grid1.Text
    Grid1.Col = 2
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
        
        Grid1.Col = 0
        bookmaster.comboAgentMaster.Text = Grid1.Text
        Grid1.Col = 1
        bookmaster.aadd1 = Grid1.Text
        Grid1.Col = 2
        bookmaster.aadd2 = Grid1.Text
        Grid1.Col = 3
        bookmaster.acity = Grid1.Text
        Grid1.Col = 4
        bookmaster.aphone = Grid1.Text
        
        Unload Me
        bookmaster.Enabled = True
        bookmaster.TextfindAgentmaster.Text = bookmaster.comboAgentMaster.Text
   
        If bookmaster.Commandmasteradd.Enabled = True Then
            If bookmaster.Commandmasteredit.Enabled = True Then
               X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
                If X = 6 Then
                        bookmaster.Agent.Enabled = True
                        bookmaster.addmaster = False
                        bookmaster.comboAgentMaster.Enabled = False
                        bookmaster.TextfindAgentmaster.Text = bookmaster.comboAgentMaster.Text
                    
                        bookmaster.Addcd.Enabled = True
                        bookmaster.Removecd.Enabled = True
                    
                        bookmaster.aadd1.Enabled = True
                        bookmaster.aadd2.Enabled = True
                        bookmaster.acity.Enabled = True
                        bookmaster.aphone.Enabled = True
                        'bookmaster.ListDis1.Enabled = True
                        'bookmaster.ListDis2.Enabled = True
                    
                        bookmaster.Commandmastersave.Enabled = True
                        bookmaster.Commandmasteredit.Enabled = True
                        bookmaster.Commandmasteradd.Enabled = True
                        bookmaster.Commandmasterdelete.Enabled = True
                        bookmaster.Agent.Enabled = True
                        bookmaster.comboAgentMaster.Enabled = True
                        SetButton bookmaster.Commandmasteradd, bookmaster.Commandmasteredit, bookmaster.Commandmastersave, bookmaster.Commandmasterdelete
                        bookmaster.comboAgentMaster.SetFocus
                        
                Else
                   
                        bookmaster.Agent.Enabled = False
                End If
            End If
        End If
    End If
    
    
    '*************   for credit not search

 
If tablabel = 15 And Trim(UCase(masterlabel)) = Trim(UCase("Creditnotefile")) Then
        Grid1.Col = 0
              
        'Creditnotefile.Frame1.Enabled = True
       
        Creditnotefile.RS.Find "Cnn = " + Trim(Val(Grid1.Text)) + "", 1
         Unload Me
        If Not Creditnotefile.RS.EOF Then
          ' X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
             ' If X = 6 Then
                       'Creditnotefile.Frame1.Enabled = False
                       Creditnotefile.Show
                       'Creditnotefile.Frame1.Enabled = False
             
         Else
          
                        Creditnotefile.Commandabandon_Click
                        Creditnotefile.Show
                        'Creditnotefile.Frame1.Enabled = False
          
        End If

    End If



If tablabel = 16 And Trim(UCase(masterlabel)) = Trim(UCase("Debitnotefile")) Then
        Grid1.Col = 0
              
        'Debitnotefile.Frame1.Enabled = True
       
        Debitnotefile.RS.Find "dnn = " + Trim(Val(Grid1.Text)) + "", 1
         Unload Me
        If Not Debitnotefile.RS.EOF Then
           'X = MsgBox("Do You Want to Edit this Record ", vbYesNo, "!!!!!!!!!!!")
            '  If X = 6 Then
                       
                       'Debitnotefile.Frame1.Enabled = False
                       'Debitnotefile.Commandedit_Click
                       Debitnotefile.Show
                       'Debitnotefile.Frame1.Enabled = False
             ' Else
              '        Debitnotefile.Frame1.Enabled = False
               '       Debitnotefile.Show
                
             ' End If
           Else
          
                        'MsgBox "Record not found"
                        Debitnotefile.Commandabandon_Click
                         'Debitnotefile.Frame1.Enabled = False
                        Debitnotefile.Show
          
          End If
          Debitnotefile.Show
    End If

End If

999:  If Err.Number = 6160 Then
         MsgBox "Record not found"
         textsearch.SetFocus
         Exit Sub
       End If

End Sub

Private Sub Grid1_GotFocus()
On Error Resume Next
For I = 0 To Data1.Recordset.Fields.Count - 1
      Grid1.Columns(I).Locked = True
Next I
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
    If tb = 0 And Trim(UCase(master)) = Trim(UCase("master")) Then
         Grid1.Columns(0).Width = 1200
         Grid1.Columns(1).Width = 3000
         Data1.RecordSource = "select * from GLEDGER where  " & stringyear & " order by gledger"
         Data1.Refresh
         Grid1.ReBind
         Grid1.Columns(5).NumberFormat = "0.00"
    End If
    
    If tb = 1 And Trim(UCase(master)) = Trim(UCase("master")) Then
          Grid1.Columns(0).Width = 3000
          Grid1.Columns(1).Width = 3000
          Data1.RecordSource = "select * from SLEDGER where   " & stringyear & " order by SUBLEDGER"
          Data1.Refresh
          Grid1.ReBind
          Grid1.Columns(3).NumberFormat = "0.00"
    End If
    
'///////////////********************////////
'//      CREDIT NOTE END PART SEARCH
'///////////////********************////////

    If tb = 3 And Trim(UCase(master)) = Trim(UCase("master")) Then
        Data1.RecordSource = "select * from CREditEND where  " & stringyear & "    order by CGENLEDGER"
        Data1.Refresh
         Grid1.Columns(2).Width = 1600
         Grid1.Columns(4).Width = 1600

    End If
    
'///////////////********************////////
'       CASH  END PART SEARCH
'///////////////********************////////

    If tb = 5 And Trim(UCase(master)) = Trim(UCase("master")) Then
         Data1.RecordSource = "select * from CASHEND  where  " & stringyear & "   order by CGENLEDGER"
         Data1.Refresh
         Grid1.Columns(2).Width = 1600
         Grid1.Columns(4).Width = 1600

    End If
    
    
    
    
'///////////////********************////////
'       invoice END PART SEARCH
'///////////////********************////////

    If tb = 2 And Trim(UCase(master)) = Trim(UCase("master")) Then
         Data1.RecordSource = "select * from invoiceend  where  FYEAR='" & main.session & "'   order by CGENLEDGER"
         Data1.Refresh
            Grid1.Columns(2).Width = 1600
            Grid1.Columns(4).Width = 1600
    End If
    
    

'///////////////********************////////
'       BOOK SEARCH
'///////////////********************////////

    If tb = 0 And Trim(UCase(master)) = Trim(UCase("bookmaster")) Then
         Grid1.Columns(0).Width = 1000
         Grid1.Columns(1).Width = 3000
         'DATA1.RecordSource = "select bookcode,bookname,quality,rate from BOOKS order by bookcode"
         Data1.RecordSource = "select * from BOOKS  where   " & stringyear & "  order by bookcode"
         Data1.Refresh
         Grid1.ReBind
         'Grid1.Columns(3).NumberFormat = "0.00"
    End If
    
'///////////////********************////////
'       discount SEARCH
'///////////////********************////////
   
    If tb = 4 And Trim(UCase(master)) = Trim(UCase("master")) Then
        Data1.RecordSource = "select * from disccats  where  FYEAR='" & main.session & "'   order by categorycode"
        Data1.Refresh
        Grid1.ReBind
        Grid1.Columns(2).NumberFormat = "0.00"
        
     End If
    
     'If tb = 5 And Trim(UCase(master)) = Trim(UCase("master")) Then
      '   DATA1.RecordSource = "select * from cbmf order by gld,sld"
      '   '  Grid1.ColWidth(0) = 3500
       '  '  Grid1.ColWidth(1) = 4500
       '  DATA1.Refresh
    'End If
    
'///////////////********************////////
'       GROUP SEARCH
'///////////////********************////////
   
    If tb = 1 And Trim(UCase(master)) = Trim(UCase("bookmaster")) Then
         Data1.RecordSource = "select * from Groups  where  " & stringyear & "  order by groupcode"
         Data1.Refresh
    End If
    
'///////////////********************////////
'       invoice SEARCH
'///////////////********************////////
    
   If tb = 13 And Trim(UCase(master)) = Trim(UCase("Purchase")) Then
      
        Data1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER,BILTYNO, NETAMOUNT from Purchasea  where " & stringyear & " order by invoiceno"
        Data1.Refresh
         Grid1.Columns(2).Width = 4000
        If Data1.Recordset.RecordCount > 0 Then
          Grid1.Columns(3).NumberFormat = "0.00"
          Grid1.Row = 0
          Grid1.Col = 0
        End If
    End If


    If tb = 11 And Trim(UCase(master)) = Trim(UCase("invoice")) Then
      
         
        Data1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER,NETAMOUNT from invoicea where  " & stringyear & "   order by invoiceno"
        Data1.Refresh
         Grid1.Columns(2).Width = 4000
        If Data1.Recordset.RecordCount > 0 Then
          Grid1.Columns(3).NumberFormat = "0.00"
          
         Grid1.Row = 0
         Grid1.Col = 0
       End If
    End If
  
    If tb = 111 And Trim(UCase(master)) = Trim(UCase("challan")) Then
      
         
        Data1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER, NETAMOUNT from casha where  " & stringyear & "   order by invoiceno"
        Data1.Refresh
         Grid1.Columns(2).Width = 4000
        If Data1.Recordset.RecordCount > 0 Then
          Grid1.Columns(3).NumberFormat = "0.00"
          
         Grid1.Row = 0
         Grid1.Col = 0
       End If
    End If
  
  
    If tb = 112 And Trim(UCase(master)) = Trim(UCase("challannew")) Then
       Data1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER, NETAMOUNT from challana  where  " & stringyear & "   order by invoiceno"
        Data1.Refresh
         Grid1.Columns(2).Width = 4000
        If Data1.Recordset.RecordCount > 0 Then
          Grid1.Columns(3).NumberFormat = "0.00"
          Grid1.Row = 0
          Grid1.Col = 0
       End If
    End If
  
  
      If tb = 12 And Trim(UCase(master)) = Trim(UCase("IssueA")) Then
      
         
        Data1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER, NETAMOUNT from IssueA  where  " & stringyear & "   order by invoiceno"
        Data1.Refresh
         Grid1.Columns(2).Width = 4000
        If Data1.Recordset.RecordCount > 0 Then
          Grid1.Columns(3).NumberFormat = "0.00"
          
         Grid1.Row = 0
         Grid1.Col = 0
       End If
    End If

  '*******************************counter sale search
  
  
    If tb = 17 And Trim(UCase(master)) = Trim(UCase("countersale")) Then
      
         
        Data1.RecordSource = "select INVOICENO as CashMemoNo,INVOICEDATE as CashMDate ,cashpartyName as Cash_Party, NETAMOUNT,SUBLEDGER from casha  where  " & stringyear & "   order by invoiceno"
        
        Data1.Refresh
        Grid1.Columns(2).Width = 4000
        
        If Data1.Recordset.RecordCount > 0 Then
        Grid1.Columns(3).NumberFormat = "0.00"
        Grid1.Row = 0
        Grid1.Col = 0
       End If
    End If







'///////////////********************////////
'       voucher SEARCH
'///////////////********************////////
   
    If tb = 12 And Trim(UCase(master)) = Trim(UCase("voucher")) Then
        Grid1.Columns(2).Width = 4000
        Data1.RecordSource = "select * from vouchers where vouchertype = '" + Voucherform.vtype + "' and  " & stringyear & "  order by voucherdate, vouchernumber"
        Data1.Refresh
    End If
    
'///////////////********************////////
'       CREDIT ITEM NOTE
'///////////////********************////////
   
    If tb = 13 And Trim(UCase(master)) = Trim(UCase("CREDITITEMNOTE")) Then
       
        Data1.RecordSource = "select INVOICENO,INVOICEDATE,SUBLEDGER,NETAMOUNT from CREDITA  where  " & stringyear & "   order by invoiceno"
        Data1.Refresh
        Grid1.Columns(2).Width = 4000
       If Data1.Recordset.RecordCount > 0 Then
       Grid1.Columns(3).NumberFormat = "0.00"
        'For I = 0 To DATA1.Recordset.RecordCount - 1
         '  Grid1.row = I
         '  Grid1.col = 3
          ' Grid1.Text = Format(Grid1.Text, "0.00")
          ' Grid1.Refresh
'         Next I
        ' Grid1.row = 0
        ' Grid1.col = 0
       End If
 
         
         
     End If
    
    
'/////////////////////////////////////////
'///////////// agent rebind/////////////////
'
'///////////////////////////////////////////

    
    If tb = 3 And Trim(UCase(master)) = Trim(UCase("bookmaster")) Then
       Grid1.Columns(0).Width = 3000
        'DATA1.RecordSource = "select Distinct Agentname from Districts where  " & stringyear & " and isnull(agentname)= false order by agentname"
        Data1.RecordSource = "select  * from Agentmaster  where " & stringyear & " order by agentname"
        
        Data1.Refresh
    Grid1.ReBind
    End If
    
    
    
    
    '****** credit note file
    
    If tb = 15 And Trim(UCase(master)) = Trim(UCase("CREDITNOTEFILE")) Then
        
        Data1.RecordSource = "select Cnn as CreditNoteNo,cnd as CreditNoteDate,PSld as SubLedger ,na as NetAmount from Cnf1a  where  " & stringyear & "   order by Cnn,cnd"
        Data1.Refresh
        Grid1.Columns(1).Width = 1000
        Grid1.Columns(2).Width = 5000
        Grid1.Columns(3).Width = 1000
        If Data1.Recordset.RecordCount > 0 Then
         Grid1.Columns(3).NumberFormat = "0.00"
         'For I = 0 To DATA1.Recordset.RecordCount - 1
          ' Grid1.row = I
           'Grid1.col = 3
           'Grid1.Columns(1).Locked =
           
           'Grid1.Columns(3).Text = Format(Grid1.Columns(3).Text, "0.00")
           'Grid1.Refresh
         'Next I
         Grid1.Row = 0
         Grid1.Col = 0
         End If
            
    
    End If
    '************ debit note file
    If tb = 16 And Trim(UCase(master)) = Trim(UCase("DebitNOTEFILE")) Then
      
        Data1.RecordSource = "select dnn as DebitNoteNo,dnd as DebitNoteDate,PSld as SubLedger ,na as NetAmount   from dnfa  where  " & stringyear & "   order by dnn,dnd"
         
        Data1.Refresh
        Grid1.Columns(1).Width = 1000
        Grid1.Columns(2).Width = 5000
        Grid1.Columns(3).Width = 1000
        If Data1.Recordset.RecordCount > 0 Then
          'Grid1.Columns(3).NumberFormat = Format(Grid1.Columns(3).Text, "0.00")
         Grid1.Columns(3).NumberFormat = "0.00"
         'For I = 0 To DATA1.Recordset.RecordCount - 1
         
          ' Grid1.row = I
           'Grid1.col = 3
          
           'DATA1.Recordset.Fields(3).Value = Format("DATA1.Recordset.Fields(3).Value", "0.00")
            'Grid1.Columns(3).Text = Format(Grid1.Text, "0.00")
             'grid1.Columns(3).Text = ""
            ' Grid1.Refresh
         'Next I
         Grid1.Row = 0
         Grid1.Col = 0
    End If
    
    End If
    
   '
    
    Me.Show
    'Grid1 = 300
    textsearch.SetFocus
End Function

Private Sub PgCntFootBeg1_GotFocus()

End Sub



Private Sub textsearch_Change()
    'Data1.Refresh
    If tablabel = 0 And Trim(UCase(masterlabel)) = UCase("master") Then
        Grid1.Col = 1
        sets = True
    End If
    If tablabel = 1 And Trim(UCase(masterlabel)) = UCase("master") Then
        Grid1.Col = 1
        sets = True
    End If
    If tablabel = 4 And Trim(UCase(masterlabel)) = UCase("master") Then
        Grid1.Col = 0
        sets = True
    End If
    
    If tablabel = 5 And Trim(UCase(masterlabel)) = UCase("master") Then
        Grid1.Col = 0
        sets = True
    End If
    'Data1.Recordset.FindFirst Trim(Data1.Recordset.Fields(grid1.col).Name) + " like '" + Trim(textsearch.Text) + "*'"
    'grid1.row = Data1.Recordset.AbsolutePosition + 1
    'grid1.SetFocus
    'SendKeys ("{LEFT}")
End Sub
Private Sub textsearch_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Data1.Refresh
        If tablabel = 0 And Trim(UCase(masterlabel)) = UCase("master") Then
            Grid1.Col = IIf(intselcol = -1, 1, intselcol)
            sets = True
        ElseIf tablabel = 1 And Trim(UCase(masterlabel)) = UCase("master") Then
            Grid1.Col = IIf(intselcol = -1, 1, intselcol)
            sets = True
        ElseIf tablabel = 4 And Trim(UCase(masterlabel)) = UCase("master") Then
            Grid1.Col = IIf(intselcol = -1, 0, intselcol)
            sets = True
        ElseIf tablabel = 11 And Trim(UCase(masterlabel)) = UCase("invoice") Then
            Grid1.Col = IIf(intselcol = -1, 0, intselcol)
            sets = True
        ElseIf tablabel = 12 And Trim(UCase(masterlabel)) = UCase("vouchers") Then
            Grid1.Col = IIf(intselcol = -1, 0, intselcol)
            sets = True
        ElseIf tablabel = 15 And Trim(UCase(masterlabel)) = UCase("CreditnoteFile") Then
            Grid1.Col = IIf(intselcol = -1, 0, intselcol)
            sets = True
        ElseIf tablabel = 17 And Trim(UCase(masterlabel)) = UCase("countersale") Then
            Grid1.Col = IIf(intselcol = -1, 0, intselcol)
            sets = True
        Else
            Grid1.Col = IIf(intselcol = -1, 0, intselcol)
        End If
        Data1.Recordset.Sort = Data1.Recordset.Fields(Grid1.Col).Name
        
       
       Dim strsearch As String
       If textsearch.Text <> "" Then
           If Data1.Recordset.RecordCount > 0 Then
              Select Case Data1.Recordset.Fields(Grid1.Col).Type
              Case 3, 5, 17: 'number
                If IsNumeric(textsearch.Text) = False Then
                MsgBox "Invalid Number."
                textsearch.SetFocus
                Exit Sub
                End If
                strsearch = " like " + Trim(textsearch.Text) + ""
              Case 202: 'text
                strsearch = " like '" + Trim(textsearch.Text) + "*'"
              Case 135: 'date
                If IsDate(textsearch.Text) = False Then
                MsgBox "Invalid Date."
                textsearch.SetFocus
                Exit Sub
                End If
                strsearch = " like '" + Trim(textsearch.Text) + "'"
              Case 11: 'boolean
              strsearch = " = " + Trim(textsearch.Text) + ""
              Case Else:
              strsearch = " = '" + Trim(textsearch.Text) + "'"
              End Select
              
              If IsNumeric(Trim(textsearch.Text)) Then
                  Data1.Recordset.Find Trim(Data1.Recordset.Fields(Grid1.Col).Name) + strsearch
              Else
                  Data1.Recordset.Find Trim(Data1.Recordset.Fields(Grid1.Col).Name) + strsearch
              End If
              
              If Data1.Recordset.AbsolutePosition > 0 Then
                  a = Data1.Recordset.AbsolutePosition
                  Data1.Recordset.AbsolutePosition = a
              End If
           End If
              Grid1.SetFocus
              For I = 0 To Data1.Recordset.Fields.Count - 1
                   Grid1.Columns(I).Locked = True
              Next I
              'Grid1.Columns(1).Locked = True
              'SendKeys ("{LEFT}")
              Grid1.SetFocus
              
           End If
    End If
End Sub

