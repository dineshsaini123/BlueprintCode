VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ViewlDisSalesRet 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   1410
   ClientWidth     =   9405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   1590
      Picture         =   "VDSalesR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   345
   End
   Begin VB.CommandButton Cprint 
      Height          =   345
      Left            =   2220
      Picture         =   "VDSalesR.frx":0451
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4590
      Width           =   345
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   4185
      Left            =   30
      TabIndex        =   2
      Top             =   90
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   7382
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   20000
      TextRTF         =   $"VDSalesR.frx":05C3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3690
      TabIndex        =   1
      Text            =   "50 %"
      Top             =   4710
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   435
      Left            =   6270
      TabIndex        =   0
      Top             =   4740
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   930
      Top             =   4470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FromPage        =   1
      PrinterDefault  =   0   'False
      ToPage          =   1
   End
End
Attribute VB_Name = "ViewlDisSalesRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLSTRING As String
Dim RS As ADODB.Recordset
Public Function rsets(ST As String, length As Integer) As String
    Dim kk As String
            kk = Trim(ST)
            If Len(kk) < length Then
                Do While Not Len(kk) = length
                    kk = " " + kk
                Loop
            End If
            If Len(kk) > length Then
                Do While Not Len(kk) = length
                    kk = Mid$(kk, 0, Len(kk) - 1)
                Loop
            End If
        rsets = kk
End Function

Private Sub Combo1_Change()
If Trim(Combo1.Text) = "50 %" Then
    r1.Font.Size = 5
End If
If Trim(Combo1.Text) = "75 %" Then
    r1.Font.Size = 8
End If
If Trim(Combo1.Text) = "100 %" Then
    r1.Font.Size = 10
End If
If Trim(Combo1.Text) = "125 %" Then
    r1.Font.Size = 12
End If
If Trim(Combo1.Text) = "150 %" Then
    r1.Font.Size = 14
End If
If Trim(Combo1.Text) = "200 %" Then
    r1.Font.Size = 18
  
End If

End Sub

Private Sub Combo1_Click()
'r1.row = 1
If Trim(Combo1.Text) = "50 %" Then
    r1.Font.Size = 5
End If
If Trim(Combo1.Text) = "75 %" Then
    r1.Font.Size = 8
End If
If Trim(Combo1.Text) = "100 %" Then
    r1.Font.Size = 10
End If
If Trim(Combo1.Text) = "200 %" Then
    r1.Font.Size = 18
End If
If Trim(Combo1.Text) = "125 %" Then
    r1.Font.Size = 12
End If
If Trim(Combo1.Text) = "150 %" Then
    r1.Font.Size = 14
End If


End Sub

Private Sub Command1_Click()
    Unload Me
    '''''MainMenu.Toolbar1.Visible = True
End Sub

Private Sub Cprint_Click()
    On Error GoTo errorcancel
    Dim frompage, topage, copies As Integer
    c1.Flags = 0
    c1.Max = 11
    c1.Flags = &H40000 Or &H4&
    c1.ShowPrinter
    frompage = c1.frompage
    topage = c1.topage
    copies = c1.copies
    If c1.Flags And &H20 Then
        c2.ShowSave
    End If
    X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(Printer.Port))
errorcancel:
End Sub

Private Sub export_Click()
    d1.ShowPrinter
    MsgBox "copies =" + Str(d1.copies)
    'd1.Copies
    'Printer.PaperSize
End Sub
Public Function printnow()
    Dim X As Long
    Dim p As Printer
    For I = 0 To Printers.Count - 1
        Set p = Printers(I)
        If Trim(p.DeviceName) = Trim(Printdlg.Comboprinters.Text) Then
            Exit For
        End If
    Next
    For I = 1 To (Printdlg.UpDown1.value)
        X = Shell("" + App.Path + "\ppp.bat " + App.Path + "\vipin.txt " & Trim(p.Port))
    Next
    Printdlg.UpDown1.value = 1
    Printdlg.Text1.Text = "1"
End Function
Private Sub Form_Load()
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30
    Cprint.Top = r1.Top + r1.Height + 30
    export.Top = r1.Top + r1.Height + 30
    Set RS = New ADODB.Recordset

    genreport
    r1.FileName = "" + App.Path + "\vipin.txt"
    r1.LoadFile (r1.FileName)
End Sub
Private Sub Form_Resize()
If Me.Width > 350 And Me.Height > 1500 Then
    r1.Width = Me.Width - 250
    r1.Height = Me.Height - 2500
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30
    Cprint.Top = r1.Top + r1.Height + 30
    export.Top = r1.Top + r1.Height + 30
End If
End Sub
Private Sub return1_Click()
    Unload Me
    '''''MainMenu.Toolbar1.Visible = True
End Sub
Function genreport1()
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim xtemp As String
    Dim trs As ADODB.Recordset
    Set trs = New ADODB.Recordset
       paperWidth = 150
        T1 = 10
        T2 = 25
        T3 = 40
        T4 = 55
        T5 = 70
        T6 = 85
        T7 = 100
        T8 = 115
        MaxLine = 50
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        main.reportdata
        main.repors.Find "reportname='" + Trim(main.reportname) + "'"
        MaxLine = main.repors!totalline
        If main.repors!comp = True Then
            paperWidth = Int(main.repors!totalcolumn * 1.75)
        Else
            paperWidth = main.repors!totalcolumn
        End If
        Open "" + App.Path + "\vipin.txt" For Output As #1
        Line = 0
header:
        Dim I As Integer
        For I = 1 To main.repors!TopMargin
            Print #1, ""
            Line = Line + 1
        Next
        If kkk.State = 1 Then
            kkk.close
        End If
        CNSetup
        kkk.Open "Select * from setup1 where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
            Print #1, Chr(27) + Chr(15) + Chr(14)
            Print #1, Tab((((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15) + LEFTM); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
            Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            Print #1, Tab(((paperWidth - Len(Trim(kkk!phone1))) / 2) + LEFTM); Trim(kkk!phone1)
            Line = Line + 3
        End If
        If trs.State = 1 Then
            trs.close
        End If
        Print #1, Chr(27) + Chr(14)
        Line = Line + 1
        trs.Open "treport", CON, adOpenKeyset, adLockReadOnly, adCmdText
        xstr = trs!Period
        Print #1, Tab(((paperWidth - (Len(Trim(trs!header)))) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); Trim(trs!header)
        Line = Line + 1
        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + Trim(xstr))
        Line = Line + 1
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Line = Line + 1
        'Print #1, Tab(LEFTM); "V.DATE"; Tab(T1 + 1 + LEFTM); "V.T."; Tab(T1 + 6 + LEFTM); "V.NUMBER"; Tab(T2 + 7 + LEFTM); "AMOUNT(Dr.)"; Tab(T3 + 12 + LEFTM); "AMOUNT(Cr.)"; Tab(t4 + 12 + LEFTM); "Ek Adh.(Dr.)"; Tab(t5 + 10 + LEFTM); "Q.Bank(Dr.)"; Tab(t6 + 10 + LEFTM); "BALANCE"
        
        Print #1, Tab(LEFTM); "V.DATE"; Tab(T1 + 1 + LEFTM); "V.T."; Tab(T1 + 6 + LEFTM); "V.NUMBER"; Tab(T2 + 11 + LEFTM); "Net. Amount"
        
        Print #1, Tab(LEFTM); repli("-", paperWidth)
        Line = Line + 1
        trs.close
        If called1 Then
            GoTo printagain1
        End If
        If rs1.State = 1 Then
            rs1.close
        End If
        
        'rs1.Open "Select "
        
        
        
        
        
        rs1.Open "select * from rptTempInDis order by VDate,vtype,Vno", CON, adOpenKeyset, adLockReadOnly, adCmdText
       'rs1.Open "select * from treport", CON, adOpenKeyset, adLockReadOnly, adCmdText

          
            Do While Not rs1.EOF
                        
              If Not rs1.BOF Then
                   Print #1, Tab(LEFTM); rs1!SUBLEDGER
                   Line = Line + 1
                   Print #1, Tab(LEFTM); Trim(rs1!disname);
                   Line = Line + 1
           
               End If
                        
                        
               
                 Print #1, Tab(LEFTM); Trim(rs1!vdate); Tab(13 + LEFTM); IIf(IsNull(rs1!vtype), "", rs1!vtype); Tab(20 + LEFTM); (Trim(rs1!vno)); Tab(36 + LEFTM); rs1!netamt
                 Line = Line + 1
                 If Line > MaxLine Then
                        called1 = True
                    
                        GoTo printfooter
printnext:
             
                        Line = 0
                        Print #1, Chr(12)
                        GoTo header
printagain1:
                        Line = 0
                        called1 = False
                End If
                
                If Not rs1.EOF Then
                    rs1.MoveNext
                End If
            Loop
printfooter:
            If Line < MaxLine Then
            Dim tempdata As ADODB.Recordset
                 Set tempdata = New ADODB.Recordset
                 tempdata.Open "Select sum(Netamt) as Nsum from rpttempindis", CON, adOpenKeyset, adLockReadOnly, adCmdText
                 Print #1, Tab(LEFTM); repli("-", paperWidth)
                 Line = Line + 1
                 Print #1, "Tolal Net Sales  :  "; Tab(36 + LEFTM); tempdata!nsum
                 Line = Line + 1
                 Print #1, Tab(LEFTM); repli("-", paperWidth)
                 Line = Line + 1
                 Do While Line < MaxLine
                    Print #1, " "
                    Line = Line + 1
                Loop
            End If
            Print #1, Tab(LEFTM); repli("-", paperWidth)
            
            If called1 Then
                GoTo printnext
            End If
        
           Print #1, ""
           Print #1, ""
           Print #1, ""
       If trs.State = 1 Then
       trs.close
       End If
       
        Close #1
End Function
Function genreport()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim rs6 As New ADODB.Recordset
Dim called1, called2 As Boolean
Dim MaxLine As Integer
Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
Dim paperWidth As Integer
Dim xtemp As String
Dim trs As ADODB.Recordset
Dim Line As Integer
Dim kkk As ADODB.Recordset
Dim Balance As Double
Dim GSUM As Double
Dim CRSUM As Double
Dim GTSUM As Double
Dim QBSUM As Double
Dim DRSUM As Double
Dim GCRSUM As Double
Dim GGTSUM As Double
Dim GQBSUM As Double
Dim GDRSUM As Double
Dim ST As String
Dim FooterYes As Boolean
Dim DName As String
DName = ""
GCRSUM = 0
GGTSUM = 0
GQBSUM = 0
GDRSUM = 0
CRSUM = 0
GTSUM = 0
QBSUM = 0
DRSUM = 0
Set trs = New ADODB.Recordset
paperWidth = 150
T1 = 10
T2 = 25
T3 = 40
T4 = 55
T5 = 70
T6 = 85
T7 = 100
T8 = 115
MaxLine = 72
called1 = False
called2 = False
FooterYes = False
Set kkk = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
main.reportdata
main.repors.Find "reportname='" + Trim(main.reportname) + "'"
MaxLine = main.repors!totalline
If main.repors!comp = True Then
    paperWidth = Int(main.repors!totalcolumn * 1.75)
Else
    paperWidth = main.repors!totalcolumn
End If
    paperWidth = 96
Open "" + App.Path + "\vipin.txt" For Output As #1
Line = 0
Pno = 1
header:
    Dim I As Integer
    For I = 1 To main.repors!TopMargin
        Print #1, ""
        Line = Line + 1
    Next
    If FooterYes = True Then
       Do While Line < 72
          Print #1, " "
          Line = Line + 1
       Loop
       Line = 0
       FooterYes = False
    End If
    If kkk.State = 1 Then kkk.close
    CNSetup
    kkk.Open "Select * from setup1 where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not kkk.BOF Then
        Print #1, ""
        Print #1, Chr(27) + Chr(77) + Chr(14)
        Print #1, Tab(84); "Page No : " & Pno
        Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15); Chr(27) + Chr(77) + Chr(14); dspace(Trim(kkk!cname))
        Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(77); dspace(Trim(kkk!add1))
        Line = Line + 5
    End If
    If trs.State = 1 Then trs.close
    
    Print #1, Tab(((paperWidth - (Len(Trim("District Wise Sales Return")))) / 2) + LEFTM); "District Wise Sales Return"
    Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Chr(27) + Chr(71); Trim("Period : " + Trim(xstr))
    Print #1, Tab(LEFTM); repli("-", paperWidth)
    Print #1, Tab(LEFTM); "V.DATE"; Tab(T1 + 1 + LEFTM); "V.T."; Tab(T1 + 6 + LEFTM); "V.NUMBER"; Tab(T2 + 4 + LEFTM); "AMOUNT(Cr.)"; Tab(T3 + 5 + LEFTM); "AMOUNT(Dr.)"; Tab(T4 + 5 + LEFTM); "AMOUNT(Dr.)"; Tab(T5 + 3 + LEFTM); "AMOUNT(Dr.)"; Tab(T6 + 4 + LEFTM); "BALANCE"
    Print #1, Tab(LEFTM); repli("-", paperWidth); Chr(27) + Chr(72)
    Line = Line + 5
    'trs.Close
    If called1 Then
       called1 = False
       GoTo printagain1
    End If
    If rs1.State = 1 Then rs1.close
    rs1.Open "Select Distinct Subleger,Distname from RPTTEMPINDIS1  where   userid = " & UId & " ORDER BY dISTNAME,Subleger ", CON, adOpenStatic
    If Not rs1.BOF Or Not rs1.EOF Then
       rs1.MoveFirst
       If DWSalesReturn.AgCombo.Text <> "" Then
          Print #1, ""
          Print #1, Tab(LEFTM); Chr(27) + Chr(14); DWSalesReturn.AgCombo.Text
          Print #1, ""
          Line = Line + 3
       End If
    End If
    DName = ""
    Do While Not rs1.EOF
            If rs2.State = 1 Then rs2.close
            If rs1.RecordCount > 0 Then
                If DName <> rs1!Distname Then
                   Print #1, ""
                   Print #1, ""
                   Print #1, Tab(LEFTM); Chr(27) + Chr(14); Trim(rs1!Distname)
                   Print #1, ""
                   Print #1, ""
                   Line = Line + 5
                End If
                Print #1, Tab(LEFTM); rs1!SUBLEGER
                Line = Line + 1
                rs2.Open "SELECT  DISTINCT  VDATE,  VNO,VTYPE  from RPTTEMPINDIS1 Where SUBLEGER = '" & rs1!SUBLEGER & "' AND  DISTNAME = '" & rs1!Distname & "' ORDER BY Vdate,Vno,VTYPE", CON, adOpenStatic, adLockOptimistic, adCmdText
                If rs2.RecordCount > 0 Then
                   rs2.MoveFirst
                   Do While Not rs2.EOF
                       If rs3.State = 1 Then rs3.close
                       rs3.Open "SELECT Sum(BNETAMT) as  GTDR from RPTTEMPINDIS1 GROUP BY GROUPCODE, SUBLEGER,DISTNAME,VDATE,VNO ,VTYPE HAVING groupcode  In ( Select groupcode from groups where group1 = 0  and group2 = 0 and " & stringyear & " ) AND  SUBLEGER = '" & rs1!SUBLEGER & "' AND  DISTNAME = '" & rs1!Distname & "' and vno = " & rs2!vno & " AND VTYPE= '" & rs2!vtype & "'  and convert(smalldatetime,vdate,103) = convert(smalldatetime,'" & rs2!vdate & "',103)", CON, adOpenStatic, adLockOptimistic, adCmdText
                       If Not rs3.BOF Then
                          s1 = 0
                          Do While Not rs3.EOF
                              s1 = Abs(s1)
                              s1 = Val(s1) + Val(rs3!GTDR)
                              rs3.MoveNext
                           Loop
                       Else
                           s1 = 0
                       End If
                       If rs4.State = 1 Then rs4.close
                       rs4.Open "SELECT Sum(BNETAMT) as QBDR from RPTTEMPINDIS1 GROUP BY GROUPCODE, SUBLEGER,DISTNAME,VDATE,VNO ,VTYPE HAVING groupcode In ( Select groupcode from groups where  group2 = 1 and " & stringyear & "  ) AND  SUBLEGER = '" & rs1!SUBLEGER & "' AND  DISTNAME = '" & rs1!Distname & "' and vno = " & rs2!vno & " AND VTYPE= '" & rs2!vtype & "' and convert(smalldatetime,vdate,103) = convert(smalldatetime,'" & rs2!vdate & "',103)", CON, adOpenStatic, adLockPessimistic, adCmdText
                       If Not rs4.BOF Then
                          s2 = 0
                          Do While Not rs4.EOF
                             s2 = Abs(s2)
                             s2 = Val(s2) + Val(rs4!QBDR)
                             rs4.MoveNext
                          Loop
                       Else
                          s2 = 0
                       End If
                       If rs6.State = 1 Then rs6.close
                       rs6.Open "SELECT Sum(BNETAMT)as AmtDr from RPTTEMPINDIS1 GROUP BY GROUPCODE, SUBLEGER,DISTNAME,VNO ,VTYPE,vdate HAVING groupcode In ( Select groupcode from groups where group1 = 1 and " & stringyear & " ) AND SUBLEGER = '" & rs1!SUBLEGER & "' AND  DISTNAME = '" & rs1!Distname & "' and vno = " & rs2!vno & " AND VTYPE= '" & rs2!vtype & "' and convert(smalldatetime,vdate,103) = convert(smalldatetime,'" & rs2!vdate & "',103)", CON, adOpenStatic, adLockOptimistic, adCmdText
                       If Not rs6.BOF Then
                           s4 = 0
                           Do While Not rs6.EOF
                               s4 = Abs(s4)
                               s4 = Val(s4) + Val(rs6!AmtDr)
                               rs6.MoveNext
                           Loop
                       Else
                           s4 = 0
                       End If
                       '********** FOR SHIFTING VNO IN CREDITB
                       
                       If rs2!vtype = "C" Then
                          ST = ""
                       Else
                          If rs5.State = 1 Then rs5.close
                          rs5.Open "SELECT Sum(BNETAMT)as AmtDr from RPTTEMPINDIS1 GROUP BY SUBLEGER,DISTNAME,VDATE,VNO ,VTYPE HAVING SUBLEGER = '" & rs1!SUBLEGER & "' AND  DISTNAME = '" & rs1!Distname & "' and vno = " & rs2!vno & " AND VTYPE= '" & rs2!vtype & "'  and convert(smalldatetime,vdate,103) = convert(smalldatetime,'" & rs2!vdate & "',103)", CON, adOpenStatic, adLockOptimistic, adCmdText
                          If Not rs5.BOF Then
                             s3 = 0
                             Do While Not rs5.EOF
                                 s3 = Abs(s3)
                                 s3 = Val(s3) + Val(rs5!AmtDr)
                                 rs5.MoveNext
                              Loop
                            Else
                                 s3 = 0
                            End If
                            ST = "   "
                        End If
                        Balance = Abs(Balance)
                        Balance = -(Balance) - s1 - s2 - s4
                        s1 = -s1
                        s2 = -s2
                        s3 = Val(s3)
                        s4 = -s4
                        'S3 = 0
                        'Print #1, Tab(LEFTM); Trim(rs2!vdate); Tab(11 + LEFTM); Trim(rs2!vtype); Tab(13 + LEFTM), ST + Trim(rs2!vno); Tab(27 + LEFTM); IIf(S1 <> 0, rsets(Trim(Format(Str(S1), "0.00")), 12), ""); Tab(45 + LEFTM); IIf(S3 <> 0, rsets(Trim(Format(Str(S3), "0.00")), 12), ""); Tab(56 + LEFTM); IIf(S4 <> 0, rsets(Trim(Format(Str(S4), "0.00")), 12), ""); Tab(70 + LEFTM); IIf(S2 <> 0, rsets(Trim(Format(Str(S2), "0.00")), 12), ""); Tab(84 + LEFTM); rsets(Trim(Format(Str(BALANCE), "0.00")), 12)
                        Print #1, Tab(LEFTM); Trim(rs2!vdate); Tab(11 + LEFTM); Trim(rs2!vtype); Tab(13 + LEFTM), ST + Trim(rs2!vno); Tab(27 + LEFTM); IIf(s3 <> 0, rsets(Trim(Format(Str(s3), "0.00")), 12), ""); Tab(43 + LEFTM); IIf(s1 <> 0, rsets(Trim(Format(Str(s1), "0.00")), 12), ""); Tab(56 + LEFTM); IIf(s4 <> 0, rsets(Trim(Format(Str(s4), "0.00")), 12), ""); Tab(70 + LEFTM); IIf(s2 <> 0, rsets(Trim(Format(Str(s2), "0.00")), 12), ""); Tab(84 + LEFTM); IIf(Balance <> 0, rsets(Trim(Format(Str(Balance), "0.00")), 12), "")
                        Line = Line + 1
                        CRSUM = CRSUM + Val(s3)
                        GTSUM = GTSUM + Val(s1)
                        QBSUM = QBSUM + Val(s2)
                        DRSUM = DRSUM + Val(s4)
                        s1 = 0
                        s2 = 0
                        s3 = 0
                        s4 = 0
                        If Line > MaxLine Then
                           called1 = True
                           FooterYes = True
                           Pno = Pno + 1
                           GoTo header
printagain1:
                           called1 = False
                        End If
                        If Not rs2.EOF Then rs2.MoveNext
                   Loop
               End If
           End If
           DName = rs1!Distname
           If Not rs1.EOF Then rs1.MoveNext
           Print #1, Tab(LEFTM); Chr(27) + Chr(71); repli("-", paperWidth)
           'Print #1, Tab(27 + LEFTM); IIf(GTSUM <> 0, rsets(Trim(Format(Str(GTSUM), "0.00")), 12), ""); Tab(56 + LEFTM); IIf(DRSUM <> 0, rsets(Trim(Format(Str(DRSUM), "0.00")), 12), ""); Tab(70 + LEFTM); IIf(QBSUM <> 0, rsets(Trim(Format(Str(QBSUM), "0.00")), 12), "")
           Print #1, Tab(43 + LEFTM); IIf(GTSUM <> 0, rsets(Trim(Format(Str(GTSUM), "0.00")), 12), ""); Tab(56 + LEFTM); IIf(DRSUM <> 0, rsets(Trim(Format(Str(DRSUM), "0.00")), 12), ""); Tab(70 + LEFTM); IIf(QBSUM <> 0, rsets(Trim(Format(Str(QBSUM), "0.00")), 12), "")
           Print #1, Tab(LEFTM); repli("-", paperWidth); Chr(27) + Chr(72)
           Line = Line + 3
           GCRSUM = GCRSUM + Abs(CRSUM)
           GGTSUM = GGTSUM + Abs(GTSUM)
           GQBSUM = GQBSUM + Abs(QBSUM)
           GDRSUM = GDRSUM + Abs(DRSUM)
           CRSUM = 0
           GTSUM = 0
           QBSUM = 0
           DRSUM = 0
           Balance = 0
           
    Loop
    If Line < MaxLine - 5 Then
       GCRSUM = -GCRSUM
       GGTSUM = -GGTSUM
       GQBSUM = -GQBSUM
       GDRSUM = -GDRSUM
       Print #1, Tab(LEFTM); Chr(27) + Chr(71); repli("-", paperWidth)
       'Print #1, "TOTAL SALES RETURN "; Tab(27 + LEFTM); IIf(GGTSUM <> 0, rsets(Trim(Format(Str(GGTSUM), "0.00")), 12), ""); Tab(45 + LEFTM); IIf(GCRSUM <> 0, rsets(Trim(Format(Str(GCRSUM), "0.00")), 12), ""); Tab(56 + LEFTM); IIf(GDRSUM <> 0, rsets(Trim(Format(Str(GDRSUM), "0.00")), 12), ""); Tab(70 + LEFTM); IIf(GQBSUM <> 0, rsets(Trim(Format(Str(GQBSUM), "0.00")), 12), "")
        'Print #1, "TOTAL SALES RETURN "; Tab(27 + LEFTM); IIf(GCRSUM <> 0, rsets(Trim(Format(Str(GCRSUM), "0.00")), 12), ""); Tab(43 + LEFTM); IIf(GGTSUM <> 0, rsets(Trim(Format(Str(GGTSUM), "0.00")), 12), ""); Tab(56 + LEFTM); IIf(GDRSUM <> 0, rsets(Trim(Format(Str(GDRSUM), "0.00")), 12), ""); Tab(70 + LEFTM); IIf(GQBSUM <> 0, rsets(Trim(Format(Str(GQBSUM), "0.00")), 12), "")
        Print #1, "TOTAL SALES RETURN "; Tab(43 + LEFTM); IIf(GGTSUM <> 0, rsets(Trim(Format(Str(GGTSUM), "0.00")), 12), ""); Tab(56 + LEFTM); IIf(GDRSUM <> 0, rsets(Trim(Format(Str(GDRSUM), "0.00")), 12), ""); Tab(70 + LEFTM); IIf(GQBSUM <> 0, rsets(Trim(Format(Str(GQBSUM), "0.00")), 12), "")
       Print #1, Tab(LEFTM); repli("-", paperWidth); Chr(27) + Chr(72)
       Line = Line + 3
       Do While Line < 72
          Print #1, " "
          Line = Line + 1
       Loop
   End If
   Close #1
End Function



        
        
        
        
        

