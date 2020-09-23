VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Graph2 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CheckBox chk_zero 
      Alignment       =   1  'Right Justify
      Caption         =   "Show 0 Values"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5970
      TabIndex        =   10
      Top             =   510
      Width           =   1275
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3195
      Left            =   0
      OleObjectBlob   =   "graph2.frx":0000
      TabIndex        =   2
      Top             =   780
      Width           =   8175
   End
   Begin MSComCtl2.DTPicker date_start 
      Height          =   285
      Left            =   3600
      TabIndex        =   8
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
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   22675459
      CurrentDate     =   38776
   End
   Begin VB.CheckBox chk_markers 
      Alignment       =   1  'Right Justify
      Caption         =   "Markers"
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
      Left            =   7335
      TabIndex        =   6
      Top             =   480
      Width           =   795
   End
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
      Height          =   285
      Left            =   1140
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   450
      Width           =   2355
   End
   Begin VB.ComboBox cmb_timeframe 
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
      ItemData        =   "graph2.frx":1981
      Left            =   0
      List            =   "graph2.frx":1991
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   450
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   4020
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2990
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   420
      ScaleHeight     =   315
      ScaleWidth      =   7305
      TabIndex        =   1
      Top             =   60
      Width           =   7300
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   60
         X2              =   7750
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
      Button_Count    =   20
      ButtonsWidth    =   31
      ButtonsHeight   =   28
      ButtonsPerRow   =   20
      HoverColor      =   -2147483641
      ToolTipBackCol  =   -2147483641
      BackGradientCol =   -2147483641
      ButtonsMode     =   5
      ButtonsSeperatorWidth=   27
      ShowSeperator   =   0   'False
      ButtonsBackColor=   -2147483641
      ButtonsGradientCol=   -2147483641
      ButtonsGradient =   3
      ButtonCaption1  =   ""
      ButtonIcon1     =   "graph2.frx":19B5
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
      ButtonCaption20 =   ""
      ButtonIcon20    =   "graph2.frx":1D07
      ButtonToolTipIcon20=   1
   End
   Begin MSComCtl2.DTPicker date_end 
      Height          =   285
      Left            =   4860
      TabIndex        =   7
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
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   22675459
      CurrentDate     =   38776
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   9
      Top             =   510
      Width           =   375
   End
End
Attribute VB_Name = "Graph2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mouse_flag As Boolean

Private Sub chk_markers_Click()

  Dim serX As Series

   If chk_markers.Value = 1 Then

      For Each serX In MSChart1.Plot.SeriesCollection
         serX.SeriesMarker.Show = True
      Next

    Else

      For Each serX In MSChart1.Plot.SeriesCollection
         serX.SeriesMarker.Show = False
      Next

   End If

End Sub

Private Sub chk_zero_Click()

   Call load_graph

End Sub

Private Sub Form_Load()

  Dim x As Integer

   date_start.Value = Date - 90
   date_end.Value = Date

   For x = 0 To (wit_num - 1)
      cmb_category.AddItem UCase((wit_cat(x)))
   Next x

   For x = 0 To (dep_num - 1)
      cmb_category.AddItem UCase((dep_cat(x)))
   Next x

   Me.Show
   Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Graph = Nothing

End Sub

Sub load_graph()

  Dim x As Integer
  Dim y
  Dim z
  Dim tmp_date As Date
  Dim hold_x As Integer
  Dim hold_frame As String
  Dim tmp_frame As String
  Dim tmp_total As Integer

   SQLString = "SELECT ENTRY_DATE, SUM(AMOUNT) FROM MAIN WHERE"
   SQLString = SQLString & " ENTRY_DATE>=" & "#" & Format(date_start.Value, "MM/DD/YY") & "#"
   SQLString = SQLString & " AND ENTRY_DATE<=" & "#" & Format(date_end.Value, "MM/DD/YY") & "#"
   SQLString = SQLString & " AND CATEGORY='" & cmb_category.Text & "'"
   SQLString = SQLString & " GROUP BY ENTRY_DATE"
   SQLString = SQLString & " ORDER BY ENTRY_DATE ASC"

   Set rs = New ADODB.Recordset
   Set rs = dB.Execute(SQLString)
   Set MSHFlexGrid1.DataSource = rs
   rs.Close
   Set rs = Nothing

   MSHFlexGrid1.Cols = 6
   MSChart1.RowCount = 0

   If chk_zero.Value = 1 Then

      tmp_date = date_start.Value

      hold_x = MSHFlexGrid1.Rows - 1

      For x = 0 To MSHFlexGrid1.Rows - 1
         y = Abs(DateDiff("d", Format(MSHFlexGrid1.TextMatrix(x, 0), "mm/dd/yyyy"), tmp_date))

         If x = 0 And Format(date_start.Value, "mm/dd/yyyy") <> Format(MSHFlexGrid1.TextMatrix(x, 0), "mm/dd/yyyy") Then
            MSHFlexGrid1.AddItem date_start.Value
         End If

         If y > 1 Then

            For z = 1 To y - 1
               MSHFlexGrid1.AddItem (tmp_date + z)
            Next z

         End If
         tmp_date = Format(MSHFlexGrid1.TextMatrix(x, 0), "mm/dd/yyyy")

         If x = hold_x And Format(date_end.Value, "mm/dd/yyyy") <> Format(MSHFlexGrid1.TextMatrix(x, 0), "mm/dd/yyyy") _
               Then
            MSHFlexGrid1.AddItem date_end.Value
         End If

      Next x

   End If

   For x = 0 To MSHFlexGrid1.Rows - 1
      MSHFlexGrid1.TextMatrix(x, 5) = Format(MSHFlexGrid1.TextMatrix(x, 0), "yyyymmdd")
      MSHFlexGrid1.TextMatrix(x, 2) = Format(MSHFlexGrid1.TextMatrix(x, 0), "YYYY")
      MSHFlexGrid1.TextMatrix(x, 3) = Format(MSHFlexGrid1.TextMatrix(x, 0), "MM")
      MSHFlexGrid1.TextMatrix(x, 4) = Format(Week(MSHFlexGrid1.TextMatrix(x, 0)), "00")

   Next x

   MSHFlexGrid1.Col = 5
   MSHFlexGrid1.Sort = 1

   For x = 0 To MSHFlexGrid1.Rows - 1

      Select Case UCase(cmb_timeframe.Text)

       Case "DAILY"

         MSChart1.RowCount = MSChart1.RowCount + 1
         MSChart1.Row = (x + 1)
         MSChart1.Data = Val(MSHFlexGrid1.TextMatrix(x, 1))
         MSChart1.RowLabel = Format(MSHFlexGrid1.TextMatrix(x, 0), "MMM-DD")

       Case "WEEKLY"
         hold_frame = MSHFlexGrid1.TextMatrix(x, 4) & MSHFlexGrid1.TextMatrix(x, 2)

         If hold_frame <> tmp_frame And tmp_frame <> "" Then
            MSChart1.RowCount = MSChart1.RowCount + 1
            MSChart1.Row = MSChart1.RowCount
            MSChart1.Data = tmp_total
            MSChart1.RowLabel = Left(tmp_frame, 2) & "/" & Right(tmp_frame, 4)

            tmp_total = 0
            tmp_total = Val(MSHFlexGrid1.TextMatrix(x, 1))
          Else
            tmp_total = tmp_total + Val(MSHFlexGrid1.TextMatrix(x, 1))
         End If

         tmp_frame = MSHFlexGrid1.TextMatrix(x, 4) & MSHFlexGrid1.TextMatrix(x, 2)

       Case "MONTHLY"

         hold_frame = MSHFlexGrid1.TextMatrix(x, 3) & "/01/" & MSHFlexGrid1.TextMatrix(x, 2)

         If hold_frame <> tmp_frame And tmp_frame <> "" Then
            MSChart1.RowCount = MSChart1.RowCount + 1
            MSChart1.Row = MSChart1.RowCount
            MSChart1.Data = tmp_total
            MSChart1.RowLabel = Format(tmp_frame, "MMM/YYYY")
            tmp_total = 0
            tmp_total = Val(MSHFlexGrid1.TextMatrix(x, 1))
          Else
            tmp_total = tmp_total + Val(MSHFlexGrid1.TextMatrix(x, 1))
         End If

         tmp_frame = MSHFlexGrid1.TextMatrix(x, 3) & "/01/" & MSHFlexGrid1.TextMatrix(x, 2)

       Case "YEARLY"

         If MSHFlexGrid1.TextMatrix(x, 2) <> tmp_frame And tmp_frame <> "" Then
            MSChart1.RowCount = MSChart1.RowCount + 1
            MSChart1.Row = MSChart1.RowCount
            MSChart1.Data = tmp_total
            MSChart1.RowLabel = tmp_frame
            tmp_total = 0
            tmp_total = Val(MSHFlexGrid1.TextMatrix(x, 1))
          Else
            tmp_total = tmp_total + Val(MSHFlexGrid1.TextMatrix(x, 1))
         End If

         tmp_frame = Val(MSHFlexGrid1.TextMatrix(x, 2))

      End Select

   Next x

   If cmb_timeframe.Text <> "DAILY" Then
      MSChart1.RowCount = MSChart1.RowCount + 1
      MSChart1.Row = MSChart1.RowCount
      MSChart1.Data = tmp_total

      Select Case UCase(cmb_timeframe.Text)

       Case "WEEKLY"
         MSChart1.RowLabel = Left(tmp_frame, 2) & "/" & Right(tmp_frame, 4)

       Case "MONTHLY"
         MSChart1.RowLabel = Format(tmp_frame, "MMM/YYYY")

       Case "YEARLY"
         MSChart1.RowLabel = tmp_frame
      End Select

   End If

End Sub

Private Sub McToolBar1_Click(ByVal ButtonIndex As Long)

   Select Case ButtonIndex
    Case 1
      mouse_flag = True
      Call load_graph

    Case 20
      Me.Hide
      Unload Me
   End Select

End Sub

Private Sub MSChart1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim intPointId As Integer
  Dim dummy As Integer
  Dim tmp_frame As String

   With MSChart1

      .AllowSeriesSelection = False

      .TwipsToChartPart x, y, dummy, dummy, intPointId, dummy, dummy

      If intPointId > 0 And mouse_flag = True Then
         MSChart1.Row = intPointId

         Select Case cmb_timeframe.Text

          Case "DAILY"
            tmp_frame = "DAY: "

          Case "WEEKLY"
            tmp_frame = "WEEK: "

          Case "MONTHLY"
            tmp_frame = "MONTH: "

          Case "YEARLY"
            tmp_frame = "YEAR: "
         End Select

         MSChart1.ToolTipText = tmp_frame & MSChart1.RowLabel & " VALUE: " & MSChart1.Data
       Else
         MSChart1.ToolTipText = ""

      End If

   End With

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

Public Function Week(ByVal inDate As Date) As Integer

   If IsDate(inDate) = False Then Exit Function

  Dim iDay As Integer
  Dim iMonth As Integer
  Dim iWeek As Integer
  Dim addDay As Integer

   iDay = 0

   Do
      iMonth = iMonth + 1

      Do
         If ((iMonth = Month(inDate)) And (iDay >= Day(inDate))) Or (iMonth > Month(inDate)) Then Week = iWeek: Exit _
            Function

         iDay = iDay + 1

         Select Case Weekday(iMonth & "/" & iDay & "/" & Year(inDate))
          Case 1: addDay = 6: iDay = iDay + addDay 'Sunday
          Case 2: addDay = 5: iDay = iDay + addDay 'Monday
          Case 3: addDay = 4: iDay = iDay + addDay 'Tuesday
          Case 4: addDay = 3: iDay = iDay + addDay 'Wednesday
          Case 5: addDay = 2: iDay = iDay + addDay 'Thursday
          Case 6: addDay = 1: iDay = iDay + addDay 'Friday
         End Select

         iWeek = iWeek + 1

         If IsDate(iMonth & "/" & (iDay + 1) & "/" & Year(inDate)) = False Then
            iDay = iDay - addDay
            'Determine the actual end of the month

            Do While IsDate(iMonth & "/" & iDay & "/" & Year(inDate)): iDay = iDay + 1: Loop
               iDay = iDay - 1
               'Determine how many days left in the week are for the next month.
               iDay = 7 - Weekday(iMonth & "/" & iDay & "/" & Year(inDate))
               Exit Do
            End If

         Loop
      Loop Until iMonth = 12

      Week = iWeek

   End Function

