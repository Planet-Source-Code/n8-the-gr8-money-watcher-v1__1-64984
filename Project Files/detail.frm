VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Detail 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CheckBox chk_auto 
      Alignment       =   1  'Right Justify
      Caption         =   "Auto-Complete On/Off"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   3240
      TabIndex        =   15
      Top             =   460
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3680
      ScaleHeight     =   315
      ScaleWidth      =   15
      TabIndex        =   14
      Top             =   2130
      Width           =   15
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Left            =   1580
      ScaleHeight     =   210
      ScaleWidth      =   3495
      TabIndex        =   11
      Top             =   2140
      Width           =   3490
      Begin VB.TextBox txt_budget 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   2130
      End
      Begin VB.ComboBox cmb_budget 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         ItemData        =   "detail.frx":0000
         Left            =   2100
         List            =   "detail.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   -30
         Width           =   1430
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   480
      ScaleHeight     =   315
      ScaleWidth      =   4155
      TabIndex        =   10
      Top             =   60
      Width           =   4155
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   5880
         Y1              =   150
         Y2              =   150
      End
   End
   Begin Money_Watcher.McToolBar McToolBar1 
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5300
      _ExtentX        =   9869
      _ExtentY        =   741
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
      Button_Count    =   30
      ButtonsWidth    =   31
      ButtonsHeight   =   28
      ButtonsPerRow   =   30
      HoverColor      =   -2147483641
      ToolTipBackCol  =   -2147483641
      BackGradientCol =   -2147483641
      ButtonsMode     =   5
      ShowSeperator   =   0   'False
      ButtonsBackColor=   -2147483641
      ButtonsGradientCol=   -2147483641
      ButtonsGradient =   3
      ButtonCaption1  =   ""
      ButtonIcon1     =   "detail.frx":0004
      ButtonToolTipIcon2=   1
      Button_Type2    =   1
      ButtonToolTipIcon3=   1
      Button_Type3    =   1
      ButtonToolTipIcon4=   1
      Button_Type4    =   1
      ButtonToolTipIcon5=   1
      Button_Type5    =   1
      ButtonToolTipIcon6=   1
      Button_Type6    =   1
      ButtonToolTipIcon7=   1
      Button_Type7    =   1
      ButtonToolTipIcon8=   1
      Button_Type8    =   1
      ButtonToolTipIcon9=   1
      Button_Type9    =   1
      ButtonToolTipIcon10=   1
      Button_Type10   =   1
      ButtonToolTipIcon11=   1
      Button_Type11   =   1
      ButtonToolTipIcon12=   1
      Button_Type12   =   1
      ButtonToolTipIcon13=   1
      Button_Type13   =   1
      ButtonToolTipIcon14=   1
      Button_Type14   =   1
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
      ButtonToolTipIcon20=   1
      Button_Type20   =   1
      ButtonToolTipIcon21=   1
      Button_Type21   =   1
      ButtonToolTipIcon22=   1
      Button_Type22   =   1
      ButtonToolTipIcon23=   1
      Button_Type23   =   1
      ButtonToolTipIcon24=   1
      Button_Type24   =   1
      ButtonToolTipIcon25=   1
      Button_Type25   =   1
      ButtonToolTipIcon26=   1
      Button_Type26   =   1
      ButtonToolTipIcon27=   1
      Button_Type27   =   1
      ButtonToolTipIcon28=   1
      Button_Type28   =   1
      ButtonToolTipIcon29=   1
      Button_Type29   =   1
      ButtonCaption30 =   ""
      ButtonIcon30    =   "detail.frx":0356
      ButtonToolTipIcon30=   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Left            =   1580
      ScaleHeight     =   210
      ScaleWidth      =   3690
      TabIndex        =   8
      Top             =   1430
      Width           =   3690
      Begin VB.ComboBox cmb_category 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         ItemData        =   "detail.frx":06A8
         Left            =   -75
         List            =   "detail.frx":06AA
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   -30
         Width           =   3600
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Left            =   1580
      ScaleHeight     =   210
      ScaleWidth      =   3690
      TabIndex        =   7
      Top             =   940
      Width           =   3690
      Begin VB.ComboBox cmb_trans 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         ItemData        =   "detail.frx":06AC
         Left            =   -75
         List            =   "detail.frx":06C2
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   -30
         Width           =   3600
      End
   End
   Begin VB.CheckBox chk_posted 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1580
      TabIndex        =   5
      Top             =   1900
      Width           =   3490
   End
   Begin VB.TextBox txt_data 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   1
      Left            =   1580
      TabIndex        =   0
      Top             =   700
      Width           =   3490
   End
   Begin VB.TextBox txt_data 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   3
      Left            =   1580
      TabIndex        =   2
      Top             =   1190
      Width           =   3490
   End
   Begin VB.TextBox txt_data 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   5
      Left            =   1580
      MaxLength       =   35
      TabIndex        =   4
      Top             =   1670
      Width           =   3490
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2040
      Left            =   0
      TabIndex        =   6
      Top             =   420
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   3598
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483645
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      GridColor       =   -2147483626
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public focus_flag As Boolean
Public edit_mode As Boolean
Private tmp_budget As Variant
Private cat_budget As String

Private Sub cmb_budget_KeyPress(KeyAscii As Integer)

   KeyAscii = AutoComplete(cmb_budget, KeyAscii, True)

End Sub

Private Sub cmb_category_KeyPress(KeyAscii As Integer)

   If chk_auto.Value = 1 Then
      KeyAscii = AutoComplete(cmb_category, KeyAscii, True)
   End If

End Sub

Private Sub cmb_category_LostFocus()

   If cmb_category.Text <> "" Then
      cat_budget = UCase(readINI("Budget", cmb_category.Text, option_file))

      If cat_budget <> "" Then
         tmp_budget = Split(cat_budget, ",")
         txt_budget.Text = Format(tmp_budget(1), "###.#0")
         cmb_budget.Text = tmp_budget(0)
      End If

   End If

End Sub

Private Sub cmb_trans_KeyPress(KeyAscii As Integer)

   If chk_auto.Value = 1 Then
      KeyAscii = AutoComplete(cmb_trans, KeyAscii, True)
   End If

End Sub

Private Sub cmb_trans_LostFocus()

  Dim x As Integer
  Dim hold_cat As String

   hold_cat = cmb_category.Text
   cmb_category.Clear
   DoEvents
   cmb_category.Text = hold_cat

   If InStr(1, UCase(Detail.cmb_trans.Text), "WITH") Then

      For x = 0 To (wit_num - 1)
         Detail.cmb_category.AddItem UCase((wit_cat(x)))
      Next x

    ElseIf InStr(1, UCase(Detail.cmb_trans.Text), "DEP") Then

      For x = 0 To (dep_num - 1)
         Detail.cmb_category.AddItem UCase((dep_cat(x)))
      Next x

    ElseIf InStr(1, UCase(Detail.cmb_trans.Text), "XFER") Then

      For x = 0 To (xfer_num - 1)
         Detail.cmb_category.AddItem UCase((xfer_cat(x)))
      Next x

   End If

End Sub

Private Sub Form_Load()

  Dim i As Integer
  Dim x As Integer

   If edit_mode = False Then focus_flag = True

   MSHFlexGrid1.ColWidth(0) = 1500
   MSHFlexGrid1.ColWidth(1) = 5000
   MSHFlexGrid1.Rows = Header.MSHFlexGrid1.Cols
   MSHFlexGrid1.Cols = 2
   MSHFlexGrid1.TextMatrix(0, 0) = "Field"
   MSHFlexGrid1.TextMatrix(0, 1) = "Value"

   For i = 1 To (Header.MSHFlexGrid1.Cols - 1)
      MSHFlexGrid1.TextMatrix(i, 0) = Header.MSHFlexGrid1.TextMatrix(0, i)

      If i = 8 Then
         MSHFlexGrid1.TextMatrix(6, 0) = Header.MSHFlexGrid1.TextMatrix(0, i)
      End If

      If edit_mode = True Then

         Select Case i
          Case 1, 3, 5
            MSHFlexGrid1.TextMatrix(i, 1) = Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.Row, i)
            txt_data(i).Text = Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.Row, i)

          Case 2
            cmb_trans.Text = Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.Row, i)

          Case 4
            cmb_category.Text = Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.Row, i)

          Case 6

            If UCase(Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.Row, i)) = "TRUE" Then
               chk_posted.Value = 1
             Else
               chk_posted.Value = 0
            End If

         End Select

      End If
   Next i

   DoEvents

   If InStr(1, UCase(Detail.cmb_trans.Text), "WITH") Then

      For x = 0 To (wit_num - 1)
         Detail.cmb_category.AddItem UCase((wit_cat(x)))
      Next x

    ElseIf InStr(1, UCase(Detail.cmb_trans.Text), "DEP") Then

      For x = 0 To (dep_num - 1)
         Detail.cmb_category.AddItem UCase((dep_cat(x)))
      Next x

   End If

   For x = 0 To (budget_num - 1)
      Detail.cmb_budget.AddItem UCase((budget_cat(x)))
   Next x

   If cmb_category.Text <> "" Then
      cat_budget = UCase(readINI("Budget", cmb_category.Text, option_file))

      If cat_budget <> "" Then
         tmp_budget = Split(cat_budget, ",")
         txt_budget.Text = Format(tmp_budget(1), "###.#0")
         cmb_budget.Text = tmp_budget(0)
      End If

   End If

   MSHFlexGrid1.Rows = Header.MSHFlexGrid1.Cols - 1
   Detail.MSHFlexGrid1.TextMatrix(1, 0) = "Entry Date"
   Detail.MSHFlexGrid1.TextMatrix(2, 0) = "Transaction"
   Detail.MSHFlexGrid1.TextMatrix(7, 0) = "Budget"
   Me.Show
   DoEvents

   If edit_mode = False Then
      txt_data(1).Text = Format(Date, "MM/DD/YYYY")
      cmb_trans.SetFocus
   End If

   Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Detail = Nothing

End Sub

Private Sub McToolBar1_Click(ByVal ButtonIndex As Long)

  Dim posted_str
  Dim write_out As String
  Dim found As Boolean
  Dim ignore_payment As Boolean
  Dim x
  Dim msg As Integer

   txt_data(5).Text = Replace(txt_data(5).Text, "'", "")
   cmb_category.Text = Replace(cmb_category.Text, "'", "")

   Select Case ButtonIndex

    Case 1

      If edit_mode = True Then
         If chk_posted.Value = 1 Then posted_str = "TRUE" Else posted_str = "FALSE"
         SQLString = ""
         SQLString = "UPDATE MAIN" & vbCrLf
         SQLString = SQLString & "SET"
         SQLString = SQLString & " ENTRY_DATE='" & Format(txt_data(1).Text, "mm/dd/yy") & "',"
         SQLString = SQLString & " TRANS='" & cmb_trans.Text & "',"
         SQLString = SQLString & " AMOUNT=" & Format(txt_data(3).Text, "####.#0") & ","
         SQLString = SQLString & " CATEGORY='" & cmb_category.Text & "', "
         SQLString = SQLString & " DESCRIPTION='" & txt_data(5).Text & "',"
         SQLString = SQLString & " POSTED=" & posted_str & vbCrLf
         SQLString = SQLString & "WHERE SEQUENCE=" & Header.MSHFlexGrid1.TextMatrix(Header.MSHFlexGrid1.RowSel, 0)

       Else
         If chk_posted.Value = 1 Then posted_str = "TRUE" Else posted_str = "FALSE"
         SQLString = ""
         SQLString = "INSERT INTO MAIN" & vbCrLf
         SQLString = SQLString & "("
         SQLString = SQLString & "ENTRY_DATE,"
         SQLString = SQLString & " TRANS,"
         SQLString = SQLString & " AMOUNT,"
         SQLString = SQLString & " CATEGORY,"
         SQLString = SQLString & " DESCRIPTION,"
         SQLString = SQLString & " POSTED)" & vbCrLf
         SQLString = SQLString & "VALUES("
         SQLString = SQLString & "'" & Format(txt_data(1).Text, "mm/dd/yyyy") & "',"
         SQLString = SQLString & " '" & cmb_trans.Text & "',"
         SQLString = SQLString & " " & txt_data(3).Text & ","
         SQLString = SQLString & " '" & cmb_category.Text & "',"
         SQLString = SQLString & " '" & txt_data(5).Text & "', "
         SQLString = SQLString & posted_str
         SQLString = SQLString & ")"

      End If

      dB.Execute UCase(SQLString)
      Header.SQL_txt.Text = UCase(SQLString)
      Set rs = New ADODB.Recordset
      Set rs = dB.Execute(DEFAULT_SQLString)
      Set Header.MSHFlexGrid1.DataSource = rs
      rs.Close
      Set rs = Nothing

      If InStr(UCase(cmb_trans.Text), "WITHDRAWAL") Then

         found = False

         For x = 0 To wit_num - 1

            If UCase(wit_cat(x)) = UCase(cmb_category.Text) Then
               found = True
               Exit For
            End If

         Next x

         If found = False Then
            If withdrawals = "" Then
               withdrawals = cmb_category.Text
             Else
               withdrawals = withdrawals & "," & cmb_category.Text
            End If

            write_out = UCase(writeINI("Defaults", "withdrawals", withdrawals, option_file))
            Call GET_transactions(1)
         End If

         ignore_payment = False

         For x = 0 To no_pay_num - 1

            If UCase(no_pay_cat(x)) = UCase(cmb_category.Text) Then
               ignore_payment = True
               Exit For
            End If

         Next x

         found = False

         For x = 0 To pay_num - 1

            If UCase(pay_cat(x)) = UCase(cmb_category.Text) Then
               found = True
               Exit For
            End If

         Next x

         If found = False And ignore_payment = False Then
            Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

            msg = MsgBox("'" & cmb_category & "'" & " does not exist as a payment." & vbCrLf & "Would you like to add?", _
               vbYesNo, "Add Payment")
            Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

            If msg = vbYes Then
               If payment = "" Then
                  payment = cmb_category.Text
                Else
                  payment = payment & "," & cmb_category.Text
               End If

               write_out = UCase(writeINI("Defaults", "payments", payment, option_file))
               Call GET_transactions(4)
             ElseIf msg = vbNo Then

               If payment_exclude = "" Then
                  payment_exclude = cmb_category.Text
                Else
                  payment_exclude = payment_exclude & "," & cmb_category.Text
               End If

               write_out = UCase(writeINI("Defaults", "payments_exclude", payment_exclude, option_file))
               Call GET_transactions(3)
            End If

         End If
      End If

      If InStr(UCase(cmb_trans.Text), "DEPOSIT") Then
         found = False

         For x = 0 To dep_num - 1

            If UCase(dep_cat(x)) = UCase(cmb_category.Text) Then
               found = True
               Exit For
            End If

         Next x

         If found = False Then
            If deposits = "" Then
               deposits = cmb_category.Text
             Else
               deposits = deposits & "," & cmb_category.Text
            End If

            write_out = UCase(writeINI("Defaults", "deposits", deposits, option_file))
            Call GET_transactions(2)
         End If

      End If

      If txt_budget.Text = "" Then cmb_budget.Text = ""

      If cmb_budget.Text <> "" Then
         write_out = UCase(writeINI("Budget", cmb_category.Text, cmb_budget.Text & "," & txt_budget.Text, option_file))
      End If

      Call Header.post_load
      DoEvents
      Me.Hide
      Unload Me
      edit_mode = False

    Case 30

      Me.Hide
      Unload Me
      edit_mode = False

   End Select

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

Private Sub txt_data_Change(Index As Integer)

   If focus_flag = True Then
      If Left(Right(txt_data(3).Text, 3), 1) = "." Then
         cmb_category.SetFocus
         focus_flag = False
      End If

   End If

End Sub

Private Sub txt_data_LostFocus(Index As Integer)

   txt_data(3).Text = Format(txt_data(3).Text, "####.#0")

End Sub

