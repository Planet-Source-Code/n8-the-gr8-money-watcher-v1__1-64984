VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Header 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Money Watcher"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ClipControls    =   0   'False
   Icon            =   "header.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   10470
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9120
      ScaleHeight     =   165
      ScaleWidth      =   885
      TabIndex        =   6
      Top             =   3900
      Width           =   915
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
   End
   Begin Money_Watcher.McToolBar McToolBar2 
      Height          =   300
      Left            =   9030
      TabIndex        =   5
      Top             =   4080
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   3
      ButtonsWidth    =   25
      ButtonsHeight   =   20
      HoverColor      =   -2147483641
      ToolTipBackCol  =   -2147483641
      BackGradientCol =   -2147483641
      ButtonsMode     =   5
      ButtonsBackColor=   -2147483641
      ButtonsGradientCol=   -2147483641
      ButtonsGradient =   3
      ButtonCaption1  =   ""
      ButtonIcon1     =   "header.frx":628A
      ButtonCaption2  =   ""
      ButtonIcon2     =   "header.frx":65DC
      ButtonToolTipIcon2=   1
      ButtonCaption3  =   ""
      ButtonIcon3     =   "header.frx":692E
      ButtonToolTipIcon3=   1
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   3180
      ScaleHeight     =   645
      ScaleWidth      =   6630
      TabIndex        =   4
      ToolTipText     =   "Double Click To Minimize"
      Top             =   0
      Width           =   6660
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   30
         X2              =   9000
         Y1              =   330
         Y2              =   330
      End
   End
   Begin Money_Watcher.McToolBar McToolBar1 
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1191
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   20
      ButtonsWidth    =   43
      ButtonsHeight   =   45
      ButtonsPerRow   =   30
      HoverColor      =   -2147483641
      TooTipStyle     =   1
      ToolTipBackCol  =   -2147483641
      BackGradientCol =   -2147483641
      ButtonsMode     =   5
      ButtonsSeperatorWidth=   20
      ShowSeperator   =   0   'False
      ButtonsBackColor=   -2147483641
      ButtonsGradientCol=   -2147483641
      ButtonsGradient =   3
      ButtonCaption1  =   ""
      ButtonIcon1     =   "header.frx":6C80
      ButtonCaption2  =   ""
      ButtonIcon2     =   "header.frx":7392
      ButtonToolTipIcon2=   1
      ButtonCaption3  =   ""
      ButtonIcon3     =   "header.frx":7AA4
      ButtonToolTipIcon3=   1
      ButtonCaption4  =   ""
      ButtonIcon4     =   "header.frx":81B6
      ButtonToolTipIcon4=   1
      ButtonCaption5  =   ""
      ButtonIcon5     =   "header.frx":88C8
      ButtonToolTipIcon5=   1
      ButtonToolTipIcon6=   1
      Button_Type6    =   1
      ButtonToolTipIcon7=   1
      Button_Type7    =   1
      ButtonToolTipIcon8=   1
      ButtonToolTipIcon9=   1
      ButtonToolTipIcon10=   1
      ButtonToolTipIcon11=   1
      ButtonToolTipIcon12=   1
      ButtonToolTipIcon13=   1
      ButtonToolTipIcon14=   1
      ButtonToolTipIcon15=   1
      Button_Type15   =   1
      ButtonToolTipIcon16=   1
      Button_Type16   =   1
      ButtonToolTipIcon17=   1
      Button_Type17   =   1
      ButtonToolTipIcon18=   1
      Button_Type18   =   1
      ButtonToolTipIcon19=   1
      Button_Type19   =   1
      ButtonCaption20 =   ""
      ButtonIcon20    =   "header.frx":8FDA
      ButtonToolTipIcon20=   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   4140
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3078
            MinWidth        =   3052
            TextSave        =   "4/6/2006"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3052
            MinWidth        =   3052
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3052
            MinWidth        =   3052
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3052
            MinWidth        =   3052
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3052
            MinWidth        =   3052
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3052
            MinWidth        =   3052
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox SQL_txt 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1395
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3060
      Width           =   10455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   690
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4154
      _Version        =   393216
      BackColor       =   -2147483639
      ForeColor       =   -2147483645
      Rows            =   8
      Cols            =   8
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      GridColor       =   -2147483637
      AllowBigSelection=   0   'False
      FillStyle       =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.Menu opt_menu 
      Caption         =   "Options"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnu_backup 
         Caption         =   "Backup Database"
      End
      Begin VB.Menu mnu_web 
         Caption         =   "Bank Website"
      End
      Begin VB.Menu mnu_datapad 
         Caption         =   "Datapad"
      End
      Begin VB.Menu mnu_graph 
         Caption         =   "Expense Pie Chart"
      End
      Begin VB.Menu mnu_chart2 
         Caption         =   "History Graph"
      End
      Begin VB.Menu mnu_budget 
         Caption         =   "Budget"
      End
      Begin VB.Menu mnu_payments 
         Caption         =   "Payments"
      End
      Begin VB.Menu mnu_reminders 
         Caption         =   "Reminders"
      End
   End
   Begin VB.Menu right_menu 
      Caption         =   "Right Click"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu right_copy 
         Caption         =   "Copy Into New"
         Index           =   2
      End
      Begin VB.Menu right_post 
         Caption         =   "Post Item"
      End
   End
   Begin VB.Menu sort_mnu 
      Caption         =   "Sort Click"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu sort_asc 
         Caption         =   "Sort Ascending"
      End
      Begin VB.Menu sort_desc 
         Caption         =   "Sort Descending"
      End
   End
End
Attribute VB_Name = "Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" ( _
      ByVal hwnd As Long, _
      ByVal lpOperation As String, _
      ByVal lpFile As String, _
      ByVal lpParameters As String, _
      ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long

Private sql_num As Integer
Private current_sql_num As Integer
Private input_file As String
Public default_sort As String
Private tmp_budget As Variant
Sub connect_DB()

   Set dB = New Connection
   dB.CursorLocation = adUseClient
   dB.Open "PROVIDER = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\" & input_file & ";"

End Sub

Sub delete_record()

  Dim msg As String

   Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
   msg = MsgBox("Delete Item: " & Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.RowSel, 0) & "?", vbYesNo, _
         "Confirm Delete")
   Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

   If msg = vbYes Then
      SQLString = "DELETE FROM MAIN" & vbCrLf
      SQLString = SQLString & "WHERE "
      SQLString = SQLString & "SEQUENCE=" & Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.RowSel, 0)
      Header.SQL_txt.Text = SQLString
      dB.Execute (SQLString)

      If Header.MSHFlexGrid1.Rows = 2 Then
         Header.MSHFlexGrid1.Clear
       Else
         Header.MSHFlexGrid1.RemoveItem (Header.MSHFlexGrid1.Row)
      End If

      Call post_load
   End If

End Sub

Private Sub Form_Load()

   If App.PrevInstance = True Then
      End
   End If

  Dim x As Integer
  Dim sql_str As String
  Dim option_file_found
  Dim msg

   option_file_found = Dir(App.Path & "\options.ini")

   If option_file_found = "" Then
      msg = MsgBox("Could not find options.ini file!" & vbCrLf & "Make sure it is located in the same directory as the" & _
         " application.", vbCritical, "Error")
      End
   End If

   option_file = App.Path & "\options.ini"
   input_file = readINI("Defaults", "input_file", option_file)
   reminders_file = readINI("Defaults", "reminders_file", option_file)
   default_sort = readINI("Defaults", "default_sort", option_file)

   For x = 1 To 999
      sql_str = readINI("SQL", Str(x), option_file)

      If sql_str = "" Then
         sql_num = x - 1
         Exit For
      End If

   Next x

   current_sql_num = 0
   Text1.Text = current_sql_num & " of " & sql_num

   Header.MSHFlexGrid1.ColWidth(0) = 700
   Header.MSHFlexGrid1.ColWidth(1) = 780
   Header.MSHFlexGrid1.ColWidth(2) = 2560
   Header.MSHFlexGrid1.ColWidth(3) = 600
   Header.MSHFlexGrid1.ColWidth(4) = 1600
   Header.MSHFlexGrid1.ColWidth(5) = 2700
   Header.MSHFlexGrid1.ColWidth(6) = 0
   Header.MSHFlexGrid1.ColWidth(7) = 700
   Header.MSHFlexGrid1.ColWidth(8) = 490

   Call connect_DB

   DEFAULT_SQLString = "SELECT * FROM MAIN ORDER BY " & default_sort
   Set rs = New ADODB.Recordset
   Set rs = dB.Execute(DEFAULT_SQLString)
   Set Header.MSHFlexGrid1.DataSource = rs
   rs.Close
   Set rs = Nothing
   DoEvents

   SQL_txt.Text = DEFAULT_SQLString
   Call GET_transactions("ALL")
   Call post_load
   Me.Show
   ShowCaption(Me) = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Header = Nothing

End Sub

Private Sub McToolBar1_Click(ByVal ButtonIndex As Long)

   Select Case ButtonIndex

    Case 1
      Detail.Show
      Detail.focus_flag = True

    Case 2
      Call delete_record

    Case 3
      Call refresh_list

    Case 4
      Search.Show

    Case 5
      PopupMenu opt_menu(1)

    Case 20
      DoEvents
      dB.Close
      Set dB = Nothing
      Me.Hide
      Unload Datapad
      Unload Detail
      Unload Graph
      Unload Graph2
      Unload Payments
      Unload Reminders
      Unload Reminders_detail
      Unload Search
      Unload Me

   End Select

End Sub

Private Sub McToolBar2_MouseDown(ByVal ButtonIndex As Long, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 y As Single)

  Dim write_out

   Select Case ButtonIndex
    Case 1
      current_sql_num = current_sql_num - 1
      If current_sql_num < 1 Then current_sql_num = sql_num
      SQL_txt.Text = UCase(Replace(readINI("SQL", Str(current_sql_num), option_file), "|", vbCrLf))

    Case 2
      write_out = UCase(writeINI("SQL", Str(sql_num + 1), Replace(SQL_txt.Text, vbCrLf, "|"), option_file))
      sql_num = sql_num + 1
      current_sql_num = sql_num

    Case 3
      current_sql_num = current_sql_num + 1
      If current_sql_num > sql_num Then current_sql_num = 1
      SQL_txt.Text = UCase(Replace(readINI("SQL", Str(current_sql_num), option_file), "|", vbCrLf))
   End Select

   Text1.Text = current_sql_num & " of " & sql_num

   DoEvents

End Sub

Private Sub mnu_backup_Click()

   dB.Close
   DoEvents
   FileCopy input_file, readINI("Defaults", "backup_input_file", option_file)
   FileCopy reminders_file, readINI("Defaults", "backup_reminders_file", option_file)
   Call connect_DB

End Sub

Private Sub mnu_budget_Click()

   Budget.Show

End Sub

Private Sub mnu_chart2_Click()

   Graph2.Show

End Sub

Private Sub mnu_datapad_Click()

   Datapad.Show

End Sub

Private Sub mnu_graph_Click()

   Graph.MSChart1.Refresh
   Graph.Show

End Sub

Private Sub mnu_payments_Click()

   Payments.Show

End Sub

Private Sub mnu_reminders_Click()

   Reminders.Show

End Sub

Private Sub mnu_web_Click()

  Dim bank_website As String
  Dim msg As String
  Dim writeout As String

   bank_website = readINI("Defaults", "bank_website", option_file)

   If bank_website <> "" Then
      ShellExecute Me.hwnd, "Open", bank_website, "", "", 3
    Else
      msg = InputBox("Enter Your Banks Website", "Bank Website Entry", "http:\\www.mybank.com")
      writeout = writeINI("Defaults", "bank_website", msg, option_file)
      bank_website = msg
      ShellExecute Me.hwnd, "Open", bank_website, "", "", 3
   End If

End Sub

Private Sub MSHFlexGrid1_DblClick()

   If Header.MSHFlexGrid1.Row > 0 Then
      Detail.edit_mode = True
      Unload Detail
      Detail.Show
   End If

End Sub

Private Sub MSHFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

   If KeyCode = "46" Then
      Call delete_record
   End If

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   tmp_col = Header.MSHFlexGrid1.MouseCol

   If Button = 2 And Header.MSHFlexGrid1.Row <> 0 And Header.MSHFlexGrid1.MouseRow <> 0 Then
      If Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.RowSel, 7) = "a" Then
         right_post.Enabled = False
       Else
         right_post.Enabled = True
      End If

      PopupMenu right_menu(2)
    ElseIf Button = 2 And Header.MSHFlexGrid1.MouseRow = 0 Then

      PopupMenu sort_mnu(3)

   End If

End Sub

Private Sub Picture2_DblClick()

   Unload Datapad
   Unload Detail
   Unload Graph
   Unload Graph2
   Unload Payments
   Unload Search
   Unload Reminders
   Unload Reminders_detail
   Unload Budget

   DoEvents

   Me.WindowState = 1

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

Sub post_load()

  Dim x As Integer
  Dim y As Integer
  Dim wfc
  Dim wfs
  Dim d2c
  Dim d2s
  Dim np
  Dim s2c
  Dim c2s
  Dim abs_wfc
  Dim abs_wfs
  Dim abs_d2c
  Dim abs_d2s
  Dim abs_np
  Dim abs_s2c
  Dim abs_c2s
  Dim reminder_count
  Dim tmp_wit As String
  Dim abs_bud
  Dim bud
  Dim start_date As Date
  Dim end_date As Date
  Dim tmp_per As Integer

   Header.MSHFlexGrid1.Cols = 9
   Header.MSHFlexGrid1.TextMatrix(0, 0) = "Sequence"
   Header.MSHFlexGrid1.TextMatrix(0, 1) = "Entry Date"
   Header.MSHFlexGrid1.TextMatrix(0, 2) = "Transaction"
   Header.MSHFlexGrid1.TextMatrix(0, 3) = "Amount"
   Header.MSHFlexGrid1.TextMatrix(0, 4) = "Category"
   Header.MSHFlexGrid1.TextMatrix(0, 5) = "Description"
   Header.MSHFlexGrid1.TextMatrix(0, 7) = "Budget"
   Header.MSHFlexGrid1.TextMatrix(0, 8) = "Posted"
   DoEvents

   For x = 1 To Header.MSHFlexGrid1.Rows - 1
      Header.MSHFlexGrid1.TextMatrix(x, 3) = Format(Header.MSHFlexGrid1.TextMatrix(x, 3), "####.#0")
      Header.MSHFlexGrid1.TextMatrix(x, 1) = Format(Header.MSHFlexGrid1.TextMatrix(x, 1), "mm/dd/yyyy")

      If Header.MSHFlexGrid1.TextMatrix(x, 6) = "True" Then
         Header.MSHFlexGrid1.Col = 8
         Header.MSHFlexGrid1.Row = x
         Header.MSHFlexGrid1.CellAlignment = 4
         Header.MSHFlexGrid1.CellFontName = "Marlett"
         Header.MSHFlexGrid1.CellFontSize = 10
         Header.MSHFlexGrid1.Text = "a"
      End If

      If InStr(1, Header.MSHFlexGrid1.TextMatrix(x, 2), "WITHDRAWAL") Then
         Header.MSHFlexGrid1.Col = 3
         Header.MSHFlexGrid1.Row = x
         Header.MSHFlexGrid1.CellForeColor = vbRed
       ElseIf InStr(1, Header.MSHFlexGrid1.TextMatrix(x, 2), "DEPOSIT") Then
         Header.MSHFlexGrid1.Col = 3
         Header.MSHFlexGrid1.Row = x
         Header.MSHFlexGrid1.CellForeColor = vbGreen
       ElseIf InStr(1, Header.MSHFlexGrid1.TextMatrix(x, 2), "XFER") Then
         Header.MSHFlexGrid1.Col = 3
         Header.MSHFlexGrid1.Row = x
         Header.MSHFlexGrid1.CellForeColor = vbBlue
      End If

   Next x

   For x = 0 To wit_num - 1

      tmp_wit = UCase(readINI("Budget", wit_cat(x), option_file))

      If tmp_wit <> "" Then

         tmp_budget = Split(tmp_wit, ",")

         Select Case tmp_budget(0)
          Case "MONTHLY"
            start_date = Format(Date, "MM") & "/01/" & Format(Date, "YY")
            end_date = Format(Date, "MM") & "/" & SetDates(start_date) & "/" & Format(Date, "YY")

          Case "WEEKLY"
            start_date = Format(DateAdd("d", 1 - Weekday(Date), Date), "MM/DD/YY")
            end_date = Format(start_date + 6, "MM/DD/YY")
         End Select

         SQLString = ""
         SQLString = "SELECT SUM(Amount) as ttl_bud FROM MAIN" & vbCrLf
         SQLString = SQLString & "WHERE TRANS = 'Withdrawal From Checkings' AND "
         SQLString = SQLString & "ENTRY_DATE>=#" & start_date & "# AND " & vbCrLf
         SQLString = SQLString & "ENTRY_DATE<=#" & end_date & "# AND " & vbCrLf
         SQLString = SQLString & "CATEGORY=" & "'" & wit_cat(x) & "'"

         bud = dB.Execute(UCase(SQLString))("ttl_bud")

         If Not IsNull(bud) Then abs_bud = Val(bud) Else abs_bud = 0

         If abs_bud > 0 And Val(tmp_budget(1)) > 0 Then

            tmp_per = (abs_bud / Val(tmp_budget(1))) * 100

            For y = 0 To Header.MSHFlexGrid1.Rows - 1

               If Header.MSHFlexGrid1.TextMatrix(y, 4) = UCase(wit_cat(x)) Then
                  Header.MSHFlexGrid1.Col = 7
                  Header.MSHFlexGrid1.Row = y
                  Header.MSHFlexGrid1.CellFontName = "Wingdings"
                  Header.MSHFlexGrid1.CellFontSize = 6

                  Header.MSHFlexGrid1.CellForeColor = vbBlue

                  Select Case tmp_per

                   Case 0 To 10
                     Header.MSHFlexGrid1.Text = "§"

                   Case 11 To 20
                     Header.MSHFlexGrid1.Text = "§§"

                   Case 21 To 30
                     Header.MSHFlexGrid1.Text = "§§§"

                   Case 31 To 40
                     Header.MSHFlexGrid1.Text = "§§§§"

                   Case 41 To 50
                     Header.MSHFlexGrid1.Text = "§§§§§"

                   Case 51 To 60
                     Header.MSHFlexGrid1.Text = "§§§§§§"

                   Case 61 To 70
                     Header.MSHFlexGrid1.Text = "§§§§§§§"

                   Case 71 To 80
                     Header.MSHFlexGrid1.Text = "§§§§§§§§"

                   Case 81 To 90
                     Header.MSHFlexGrid1.Text = "§§§§§§§§§"

                   Case 91 To 100
                     Header.MSHFlexGrid1.Text = "§§§§§§§§§§"

                   Case Is > 100
                     Header.MSHFlexGrid1.CellForeColor = vbRed
                     Header.MSHFlexGrid1.Text = "§§§§§§§§§§"
                  End Select

                  Exit For

               End If
            Next y

         End If

      End If

   Next x

   wfc = dB.Execute("SELECT SUM(Amount) as ttl_wfc FROM MAIN where TRANS = 'Withdrawal From Checkings' and POSTED=TRUE")("ttl_wfc")
   wfs = dB.Execute("SELECT SUM(Amount) as ttl_wfs FROM MAIN where TRANS = 'Withdrawal From Savings'")("ttl_wfs")
   d2c = dB.Execute("SELECT SUM(Amount) as ttl_d2c FROM MAIN where TRANS = 'Deposit To Checkings'")("ttl_d2c")
   d2s = dB.Execute("SELECT SUM(Amount) as ttl_d2s FROM MAIN where TRANS = 'Deposit To Savings'")("ttl_d2s")
   np = dB.Execute("SELECT SUM(Amount) as ttl_np FROM MAIN where POSTED = FALSE")("ttl_np")
   s2c = dB.Execute("SELECT SUM(Amount) as ttl_s2c FROM MAIN where TRANS = 'XFER SAVINGS TO CHECKINGS'")("ttl_s2c")
   c2s = dB.Execute("SELECT SUM(Amount) as ttl_c2s FROM MAIN where TRANS = 'XFER CHECKINGS TO SAVINGS'")("ttl_c2s")

   If Not IsNull(wfc) Then abs_wfc = Val(wfc) Else abs_wfc = 0
   If Not IsNull(wfs) Then abs_wfs = Val(wfs) Else abs_wfs = 0
   If Not IsNull(d2c) Then abs_d2c = Val(d2c) Else abs_d2c = 0
   If Not IsNull(d2s) Then abs_d2s = Val(d2s) Else abs_d2s = 0
   If Not IsNull(np) Then abs_np = Val(np) Else abs_np = 0
   If Not IsNull(s2c) Then abs_s2c = Val(s2c) Else abs_s2c = 0
   If Not IsNull(c2s) Then abs_c2s = Val(c2s) Else abs_c2s = 0

   Set rem_dB2 = New Connection
   rem_dB2.CursorLocation = adUseClient
   rem_dB2.Open "PROVIDER = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\" & reminders_file & ";"
   reminder_count = rem_dB2.Execute("SELECT COUNT(*) as msgs FROM REMINDERS where RREAD=FALSE and rDATE<=" & "#" & Format(Date, "MM/DD/YY") & "#")("msgs")
   rem_dB2.Close

   StatusBar1.Panels(2).Text = (Header.MSHFlexGrid1.Rows - 1) & " Records"
   StatusBar1.Panels(6).Text = reminder_count & " Reminder(s)"
   StatusBar1.Panels(4).Text = "Checkings: " & FormatNumber((abs_d2c + abs_s2c) - (abs_wfc + abs_c2s), 2, , vbTrue)
   StatusBar1.Panels(5).Text = "Savings: " & FormatNumber((abs_d2s + abs_c2s) - (abs_wfs + abs_s2c), 2, , vbTrue)
   StatusBar1.Panels(3).Text = "Available: " & FormatNumber((abs_d2c - abs_wfc) + (abs_d2s - abs_wfs) - abs_np, 2, , vbTrue)
   Header.MSHFlexGrid1.Refresh
   DoEvents

End Sub

Sub refresh_list()

   Header.SQL_txt.Text = DEFAULT_SQLString
   Set rs = New ADODB.Recordset
   Set rs = dB.Execute(DEFAULT_SQLString)
   Set Header.MSHFlexGrid1.DataSource = rs
   rs.Close
   Set rs = Nothing
   Call post_load

End Sub

Private Sub right_copy_Click(Index As Integer)

   Detail.edit_mode = True
   Unload Detail
   Detail.Show
   DoEvents
   Detail.edit_mode = False
   Detail.txt_data(1).Text = Date

End Sub

Private Sub right_post_Click()

   SQLString = ""
   SQLString = "UPDATE MAIN" & vbCrLf
   SQLString = SQLString & "SET"
   SQLString = SQLString & " POSTED=TRUE" & vbCrLf
   SQLString = SQLString & "WHERE SEQUENCE=" & Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.RowSel, 0)

   dB.Execute UCase(SQLString)
   Header.SQL_txt.Text = UCase(SQLString)

   Header.MSHFlexGrid1.Col = 8
   Header.MSHFlexGrid1.CellFontName = "Marlett"
   Header.MSHFlexGrid1.CellFontSize = 10
   Header.MSHFlexGrid1.Text = "a"
   Header.MSHFlexGrid1.CellAlignment = 4

   Header.MSHFlexGrid1.Col = 6
   Header.MSHFlexGrid1.Text = "TRUE"

   Call post_load

End Sub

Private Sub sort_asc_Click()

  Dim sort_field As String

   If UCase(Header.MSHFlexGrid1.TextMatrix(0, tmp_col)) = "ENTRY DATE" Then
      sort_field = "ENTRY_DATE"
    ElseIf UCase(Header.MSHFlexGrid1.TextMatrix(0, tmp_col)) = "TRANSACTION" Then
      sort_field = "TRANS"
    Else
      sort_field = UCase(Header.MSHFlexGrid1.TextMatrix(0, tmp_col))
   End If

   Set rs = New ADODB.Recordset
   Set rs = dB.Execute("SELECT * FROM MAIN ORDER BY " & sort_field & " ASC")
   Set Header.MSHFlexGrid1.DataSource = rs
   rs.Close
   Set rs = Nothing
   Call post_load

End Sub

Private Sub sort_desc_Click()

  Dim sort_field As String

   If UCase(Header.MSHFlexGrid1.TextMatrix(0, tmp_col)) = "ENTRY DATE" Then
      sort_field = "ENTRY_DATE"
    ElseIf UCase(Header.MSHFlexGrid1.TextMatrix(0, tmp_col)) = "TRANSACTION" Then
      sort_field = "TRANS"
    Else
      sort_field = UCase(Header.MSHFlexGrid1.TextMatrix(0, tmp_col))
   End If

   Set rs = New ADODB.Recordset
   Set rs = dB.Execute("SELECT * FROM MAIN ORDER BY " & sort_field & " DESC")
   Set Header.MSHFlexGrid1.DataSource = rs
   rs.Close
   Set rs = Nothing
   Call post_load

End Sub

Private Sub SQL_txt_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 And SQL_txt.Text <> "" Then
      KeyAscii = 0
      Dim ret_val As New ADODB.Recordset
      Dim x As Integer
      Dim parm As String

      On Error GoTo ohno:

      If InStr(1, LCase(SQL_txt.Text), "select") > 0 Then
         If Left(SQL_txt.Text, 1) = "=" Then
            Set ret_val = New ADODB.Recordset
            Set ret_val = dB.Execute(Mid(SQL_txt.Text, 2, Len(SQL_txt.Text)))

            For x = 1 To ret_val.Fields.Count
               parm = "X" & Trim(Str(x))
               SQL_txt.Text = SQL_txt.Text & vbCrLf & "Result " & parm & ": " & ret_val(parm)
            Next x

            ret_val.Close
            Set ret_val = Nothing
          Else
            Set rs = dB.Execute(SQL_txt.Text)
            Set Header.MSHFlexGrid1.DataSource = rs
            rs.Close
            Set rs = Nothing
         End If

       Else
         dB.Execute SQL_txt.Text
         Set rs = New ADODB.Recordset
         Set rs = dB.Execute(DEFAULT_SQLString)
         Set Header.MSHFlexGrid1.DataSource = rs
         rs.Close
         Set rs = Nothing
      End If

      Call post_load
      Exit Sub
      ohno:
      SQL_txt.Text = ""
      SQL_txt.Text = Err.Description

   End If

End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)

   If Panel.Index = 6 Then
      Reminders.Show
   End If

End Sub

