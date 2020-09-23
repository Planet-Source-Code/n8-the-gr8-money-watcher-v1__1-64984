VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Reminders 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3195
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   5636
      _Version        =   393216
      ForeColor       =   -2147483645
      Cols            =   6
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      GridColor       =   -2147483633
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   6
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   780
      ScaleHeight     =   315
      ScaleWidth      =   4965
      TabIndex        =   1
      Top             =   60
      Width           =   4965
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   30
         X2              =   7720
         Y1              =   150
         Y2              =   150
      End
   End
   Begin Money_Watcher.McToolBar McToolBar1 
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
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
      ButtonsWidth    =   24
      ButtonsHeight   =   28
      ButtonsPerRow   =   30
      HoverColor      =   -2147483641
      ToolTipBackCol  =   -2147483641
      BackGradientCol =   -2147483641
      ButtonsMode     =   5
      ButtonsSeperatorWidth=   2
      ShowSeperator   =   0   'False
      ButtonsBackColor=   -2147483641
      ButtonsGradientCol=   -2147483641
      ButtonsGradient =   3
      ButtonCaption1  =   ""
      ButtonIcon1     =   "reminders.frx":0000
      Button_Type1    =   1
      ButtonCaption2  =   ""
      ButtonIcon2     =   "reminders.frx":0352
      ButtonToolTipIcon2=   1
      ButtonCaption3  =   ""
      ButtonIcon3     =   "reminders.frx":06A4
      ButtonToolTipIcon3=   1
      ButtonToolTipIcon4=   1
      Button_Type4    =   1
      ButtonToolTipIcon5=   1
      ButtonToolTipIcon6=   1
      ButtonToolTipIcon7=   1
      ButtonToolTipIcon8=   1
      ButtonToolTipIcon9=   1
      ButtonToolTipIcon10=   1
      ButtonToolTipIcon11=   1
      ButtonToolTipIcon12=   1
      ButtonToolTipIcon13=   1
      ButtonToolTipIcon14=   1
      ButtonToolTipIcon15=   1
      ButtonToolTipIcon16=   1
      ButtonToolTipIcon17=   1
      Button_Type17   =   1
      ButtonToolTipIcon18=   1
      Button_Type18   =   1
      ButtonToolTipIcon19=   1
      Button_Type19   =   1
      ButtonCaption20 =   ""
      ButtonIcon20    =   "reminders.frx":09F6
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
      ButtonCaption30 =   ""
      ButtonIcon30    =   "reminders.frx":0D48
      ButtonToolTipIcon30=   1
   End
End
Attribute VB_Name = "Reminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub connect_DB()

   Set rem_dB = New Connection
   rem_dB.CursorLocation = adUseClient
   rem_dB.Open "PROVIDER = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\" & reminders_file & ";"

End Sub

Private Sub Form_Load()

   MSHFlexGrid1.ColWidth(0) = 0
   MSHFlexGrid1.ColWidth(1) = 800
   MSHFlexGrid1.ColWidth(2) = 900
   MSHFlexGrid1.ColWidth(3) = 3520
   MSHFlexGrid1.ColWidth(4) = 0
   MSHFlexGrid1.ColWidth(5) = 0
   MSHFlexGrid1.ColWidth(6) = 0
   MSHFlexGrid1.ColWidth(7) = 685

   Call connect_DB

   Set rs = New ADODB.Recordset
   Set rs = rem_dB.Execute("SELECT * FROM REMINDERS WHERE RDATE>=#" & Format(Date, "MM/DD/YY") & "#" & " OR RREAD=FALSE" & _
         " ORDER BY RDATE ASC")
   Set MSHFlexGrid1.DataSource = rs
   rs.Close
   Set rs = Nothing

   Call post_load
   Me.Show
   Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Reminders = Nothing

End Sub

Private Sub McToolBar1_Click(ByVal ButtonIndex As Long)

  Dim msg As String

   Select Case ButtonIndex
    Case 2
      Reminders_detail.Show

    Case 3
      SQLString = ""

      If Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 3) <> "ONE TIME" Then
         Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
         msg = MsgBox("Would you like to delete all " & Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, _
            3) & " " & Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 2) & " reminders?", vbYesNoCancel, _
            "Delete Reminder")
         Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
       Else
         msg = vbNo
      End If

      If msg = vbYes Then
         SQLString = "DELETE FROM REMINDERS" & vbCrLf
         SQLString = SQLString & "WHERE "
         SQLString = SQLString & "rMESSAGE='" & Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 3) & "'"
         rem_dB.Execute (SQLString)
         DoEvents
         Set rs = New ADODB.Recordset
         Set rs = rem_dB.Execute("SELECT * FROM REMINDERS WHERE RDATE>=#" & Format(Date, "MM/DD/YY") & "#" & " OR" & _
            " RREAD=FALSE ORDER BY RDATE ASC")
         Set MSHFlexGrid1.DataSource = rs
         rs.Close
         Set rs = Nothing

         Call post_load

       ElseIf msg = vbNo Then
         SQLString = "DELETE FROM REMINDERS" & vbCrLf
         SQLString = SQLString & "WHERE "
         SQLString = SQLString & "SEQUENCE=" & Reminders.MSHFlexGrid1.TextMatrix(Reminders.MSHFlexGrid1.Row, 0)
         rem_dB.Execute (SQLString)
         DoEvents
         Set rs = New ADODB.Recordset
         Set rs = rem_dB.Execute("SELECT * FROM REMINDERS WHERE RDATE>=#" & Format(Date, "MM/DD/YY") & "#" & " OR" & _
            " RREAD=FALSE ORDER BY RDATE ASC")
         Set MSHFlexGrid1.DataSource = rs
         rs.Close
         Set rs = Nothing

         Call post_load

      End If

    Case 30
      rem_dB.Close
      Set rem_dB = Nothing
      Reminders_detail.Hide
      Unload Reminders_detail
      Me.Hide
      Unload Me
   End Select

End Sub

Private Sub MSHFlexGrid1_DblClick()

   If MSHFlexGrid1.Row > 0 Then
      rem_edit_mode = True
      Unload Reminders_detail
      Reminders_detail.Show
   End If

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

Sub post_load()

  Dim x As Integer
  Dim reminder_date As Date

   MSHFlexGrid1.Cols = 8
   MSHFlexGrid1.TextMatrix(0, 1) = "Date"
   MSHFlexGrid1.TextMatrix(0, 2) = "Time"
   MSHFlexGrid1.TextMatrix(0, 3) = "Message"
   MSHFlexGrid1.TextMatrix(0, 7) = "Done"

   For x = 1 To MSHFlexGrid1.Rows - 1

      reminder_date = MSHFlexGrid1.TextMatrix(x, 1)

      MSHFlexGrid1.TextMatrix(x, 1) = Format(MSHFlexGrid1.TextMatrix(x, 1), "mm/dd/yyyy")
      MSHFlexGrid1.TextMatrix(x, 2) = Format(MSHFlexGrid1.TextMatrix(x, 2), "h:mm AMPM")

      If MSHFlexGrid1.TextMatrix(x, 5) = "True" Then
         MSHFlexGrid1.Col = 7
         MSHFlexGrid1.Row = x
         MSHFlexGrid1.CellFontName = "Marlett"
         MSHFlexGrid1.CellFontSize = 10
         MSHFlexGrid1.Text = "a"
         MSHFlexGrid1.CellAlignment = 4
      End If

      If (reminder_date <= Date And MSHFlexGrid1.TextMatrix(x, 5) = "False") Or (reminder_date = Date) Then
         MSHFlexGrid1.Row = x
         MSHFlexGrid1.Col = 1
         MSHFlexGrid1.CellForeColor = vbBlue
      End If

   Next x
   Reminders.MSHFlexGrid1.Refresh

End Sub

