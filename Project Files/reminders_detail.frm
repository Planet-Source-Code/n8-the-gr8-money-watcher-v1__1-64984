VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Reminders_detail 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CheckBox chk_read 
      Alignment       =   1  'Right Justify
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4420
      TabIndex        =   4
      Top             =   480
      Width           =   660
   End
   Begin VB.ComboBox cmb_cycle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "reminders_detail.frx":0000
      Left            =   1140
      List            =   "reminders_detail.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   450
      Width           =   1155
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   0
      Width           =   5595
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
      ButtonIcon1     =   "reminders_detail.frx":0041
      ButtonCaption2  =   ""
      ButtonIcon2     =   "reminders_detail.frx":0393
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
      ButtonIcon30    =   "reminders_detail.frx":06E5
      ButtonToolTipIcon30=   1
   End
   Begin VB.TextBox msg_txt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   750
      Width           =   5100
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   450
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   55836675
      CurrentDate     =   38762
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   450
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "h:mm tt"
      Format          =   55836675
      UpDown          =   -1  'True
      CurrentDate     =   38762.8333333333
   End
End
Attribute VB_Name = "Reminders_detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private use_outlook_calendar As String
Private calendar_name As String

Sub add_outlook_appt()

   On Error GoTo ohno:

  Dim oApp As Outlook.Application
  Dim oNS As Outlook.NameSpace
  Dim oAppt As Outlook.AppointmentItem
  Dim oApptx As Outlook.AppointmentItem
  Dim myFolder
  Dim itmsAppoint As Outlook.Items
  Dim i As Integer
  Dim x As Integer
  Dim field1 As String
  Dim field2 As String
  Dim field3 As String
  Dim tmp_date As Date

   Set oApp = New Outlook.Application
   Set oNS = oApp.GetNamespace("mapi")
   Set oAppt = oApp.CreateItem(Outlook.OlItemType.olAppointmentItem)

   Set myFolder = oNS.GetDefaultFolder(9)
   Set myFolder = myFolder.Folders(calendar_name)
   Set itmsAppoint = myFolder.Items

   For i = itmsAppoint.Count To 1 Step -1
      Set oApptx = itmsAppoint.Item(i)
      oApptx.Delete
   Next i

   For x = 1 To Reminders.MSHFlexGrid1.Rows - 1

      field1 = Reminders.MSHFlexGrid1.TextMatrix(x, 1)
      field2 = Reminders.MSHFlexGrid1.TextMatrix(x, 2)
      field3 = Reminders.MSHFlexGrid1.TextMatrix(x, 3)
      tmp_date = field1

      If tmp_date >= Date Then

         Set oApp = New Outlook.Application
         Set oNS = oApp.GetNamespace("mapi")
         Set oAppt = oApp.CreateItem(Outlook.OlItemType.olAppointmentItem)

         oAppt.Subject = field3
         oAppt.Body = ""
         oAppt.Location = ""

         oAppt.Start = CDate(field1 & " " & field2)

         oAppt.ReminderSet = True
         oAppt.Duration = 60
         oAppt.ReminderMinutesBeforeStart = 120
         oAppt.BusyStatus = Outlook.OlBusyStatus.olBusy
         oAppt.IsOnlineMeeting = False
         oAppt.Save
         oAppt.Move (myFolder)

         oNS.Logoff
         Set oApp = Nothing
         Set oNS = Nothing
         Set oAppt = Nothing
      End If

   Next x
   Exit Sub
ohno:

End Sub

Private Sub cmb_cycle_KeyPress(KeyAscii As Integer)

   KeyAscii = AutoComplete(cmb_cycle, KeyAscii, True)

End Sub

Private Sub Form_Load()

   Me.Show

   use_outlook_calendar = readINI("Defaults", "use_outlook_calendar", option_file)
   calendar_name = readINI("Defaults", "calendar_name", option_file)

   If rem_edit_mode = True Then
      DTPicker1.Value = Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 1)

      If Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 2) <> "" Then
         DTPicker2.Value = Format(Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 2), "h:mm AMPM")
       Else
         DTPicker2.Value = "12:00 AM"
      End If

      msg_txt.Text = Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 3)
      cmb_cycle.Text = Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 4)

      If UCase(Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 5)) = "TRUE" Then
         chk_read.Value = 1
       Else
         chk_read.Value = 0
      End If

      cmb_cycle.Enabled = False

    Else
      cmb_cycle.Text = "ONE TIME"
      DTPicker1.Value = Date
      DTPicker2.Value = "12:00 AM"
      msg_txt.SetFocus
      cmb_cycle.Enabled = True

   End If

   Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Reminders_detail = Nothing

End Sub

Private Sub McToolBar1_Click(ByVal ButtonIndex As Long)

  Dim repeat_counter As Long
  Dim x As Integer
  Dim read_str As String
  Dim t_day
  Dim t_month
  Dim t_year As String

   msg_txt.Text = Replace(msg_txt.Text, "'", "")

   Select Case ButtonIndex

    Case 1
      If chk_read.Value = 1 Then read_str = "TRUE" Else read_str = "FALSE"

      If rem_edit_mode = True Then
         SQLString = ""
         SQLString = "UPDATE REMINDERS" & vbCrLf
         SQLString = SQLString & "SET"
         SQLString = SQLString & " rDATE='" & Format(DTPicker1.Value, "mm/dd/yy") & "',"
         SQLString = SQLString & " rTIME='" & Format(DTPicker2.Value, "hh:mm") & "',"
         SQLString = SQLString & " rCycle='" & cmb_cycle.Text & "',"
         SQLString = SQLString & " rMESSAGE='" & msg_txt.Text & "',"
         SQLString = SQLString & " rRead=" & read_str
         SQLString = SQLString & " WHERE SEQUENCE=" & Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.RowSel, _
            0)
         rem_dB.Execute UCase(SQLString)

       Else

         If InStr(1, cmb_cycle.Text, "WEEKLY") > 0 Or InStr(1, cmb_cycle.Text, "YEARLY") > 0 Or InStr(1, cmb_cycle.Text, _
               "DAILY") > 0 Or InStr(1, cmb_cycle.Text, "MONTHLY") > 0 Then
            Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            Call SetWindowPos(Reminders.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

            repeat_counter = InputBox("Enter The Number Of Occurences For This Reminder", "Reminders", "0")
            Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            Call SetWindowPos(Reminders.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
         End If

         If repeat_counter = 0 Then repeat_counter = 1

         For x = 1 To repeat_counter

            SQLString = ""
            SQLString = "INSERT INTO REMINDERS" & vbCrLf
            SQLString = SQLString & "("
            SQLString = SQLString & "rDATE,"
            SQLString = SQLString & "rTIME,"
            SQLString = SQLString & "rCycle,"
            SQLString = SQLString & " rMessage)"
            SQLString = SQLString & "VALUES("
            SQLString = SQLString & "'" & Format(DTPicker1.Value, "mm/dd/yyyy") & "',"
            SQLString = SQLString & "'" & Format(DTPicker2.Value, "hh:mm") & "',"
            SQLString = SQLString & " '" & cmb_cycle.Text & "',"
            SQLString = SQLString & " '" & msg_txt.Text & "'"
            SQLString = SQLString & ")"
            rem_dB.Execute UCase(SQLString)

            t_day = Format(DTPicker1.Value, "DD")
            t_month = Format(DTPicker1.Value, "MM")
            t_year = Format(DTPicker1.Value, "YYYY")

            Select Case UCase(cmb_cycle.Text)

             Case "DAILY"
               DTPicker1.Value = DTPicker1.Value + 1

             Case "MONTHLY"
               t_month = t_month + 1

               If t_month > 12 Then
                  t_month = "01"
                  t_year = t_year + 1
               End If

               DTPicker1.Value = t_day & "/" & t_month & "/" & t_year

             Case "WEEKLY"
               DTPicker1.Value = DTPicker1.Value + 7

             Case "YEARLY"
               t_year = t_year + 1
               DTPicker1.Value = t_day & "/" & t_month & "/" & t_year

            End Select

         Next x

      End If

      Set rs = New ADODB.Recordset
      Set rs = rem_dB.Execute("SELECT * FROM REMINDERS WHERE RDATE>=#" & Format(Date, "MM/DD/YY") & "#" & " OR" & _
            " RREAD=FALSE ORDER BY RDATE ASC")
      Set Reminders.MSHFlexGrid1.DataSource = rs
      rs.Close
      Set rs = Nothing
      Call Reminders.post_load
      Call Header.post_load
      DoEvents
      If UCase(use_outlook_calendar) = "YES" Then
        Call add_outlook_appt
      End If
      rem_edit_mode = False
      Me.Hide
      Unload Me

    Case 30

      rem_edit_mode = False
      Me.Hide
      Unload Me

   End Select

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

