VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Budget 
   BackColor       =   &H80000008&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
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
      Left            =   30
      ScaleHeight     =   315
      ScaleWidth      =   4290
      TabIndex        =   2
      Top             =   60
      Width           =   4285
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   5900
         Y1              =   150
         Y2              =   150
      End
   End
   Begin Money_Watcher.McToolBar McToolBar1 
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4830
      _ExtentX        =   8520
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
      Button_Type1    =   1
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
      ButtonIcon30    =   "budget.frx":0000
      ButtonToolTipIcon30=   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2385
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   4207
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483645
      Cols            =   5
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      BackColorBkg    =   -2147483639
      GridColor       =   -2147483626
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
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
      _Band(0).Cols   =   5
   End
End
Attribute VB_Name = "Budget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  Dim tmp_wit As String
  Dim bud2
  Dim bud
  Dim start_date As Date
  Dim end_date As Date
  Dim tmp_budget As Variant
  Dim tmp_timeframe As Variant
  Dim c_row As Integer
  Dim x
  Dim y
  Dim col2_tot As Double
  Dim col3_tot As Double
  Dim col4_tot As Double

   MSHFlexGrid1.Cols = 5
   MSHFlexGrid1.ColWidth(0) = 800
   MSHFlexGrid1.ColWidth(1) = 1410
   MSHFlexGrid1.ColWidth(2) = 770
   MSHFlexGrid1.ColWidth(3) = 770
   MSHFlexGrid1.ColWidth(4) = 770
   MSHFlexGrid1.TextMatrix(0, 0) = "Timeframe"
   MSHFlexGrid1.TextMatrix(0, 1) = "Category"
   MSHFlexGrid1.TextMatrix(0, 2) = "Budgeted"
   MSHFlexGrid1.TextMatrix(0, 3) = "Actual"
   MSHFlexGrid1.TextMatrix(0, 4) = "Difference"

   MSHFlexGrid1.Row = 0
   MSHFlexGrid1.Col = 4
   MSHFlexGrid1.CellAlignment = 7
   MSHFlexGrid1.Col = 3
   MSHFlexGrid1.CellAlignment = 7
   MSHFlexGrid1.Col = 2
   MSHFlexGrid1.CellAlignment = 7

   c_row = 0

   For x = 0 To budget_num - 1

      If x > 0 Then
         c_row = c_row + 1
         MSHFlexGrid1.Rows = c_row + 1
         MSHFlexGrid1.TextMatrix(c_row, 1) = "TOTALS:"
         MSHFlexGrid1.Col = 1
         MSHFlexGrid1.Row = c_row
         MSHFlexGrid1.CellAlignment = 7
         MSHFlexGrid1.TextMatrix(c_row, 2) = Format(col2_tot, "####.#0")
         MSHFlexGrid1.TextMatrix(c_row, 3) = Format(col3_tot, "####.#0")
         MSHFlexGrid1.TextMatrix(c_row, 4) = Format(col4_tot, "####.#0")
         MSHFlexGrid1.Col = 4
         MSHFlexGrid1.Row = c_row

         If Val(MSHFlexGrid1.TextMatrix(c_row, 4)) > 0 Then
            MSHFlexGrid1.CellForeColor = vbRed
          Else
            MSHFlexGrid1.CellForeColor = vbGreen
         End If

      End If

      c_row = c_row + 1
      MSHFlexGrid1.Rows = c_row + 1
      MSHFlexGrid1.TextMatrix(c_row, 0) = UCase(Split(budget_timeframes, ",")(x))

      For y = 0 To wit_num - 1
         tmp_wit = UCase(readINI("Budget", wit_cat(y), option_file))

         If tmp_wit <> "" Then

            tmp_budget = Split(tmp_wit, ",")
            tmp_timeframe = Split(budget_timeframes, ",")

            If UCase(tmp_budget(0)) = UCase(tmp_timeframe(x)) Then

               c_row = c_row + 1
               MSHFlexGrid1.Rows = c_row + 1
               MSHFlexGrid1.TextMatrix(c_row, 1) = UCase(wit_cat(y))
               MSHFlexGrid1.TextMatrix(c_row, 2) = Format(Split(tmp_wit, ",")(1), "####.#0")
               col2_tot = col2_tot + Val(Split(tmp_wit, ",")(1))

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
               SQLString = SQLString & "CATEGORY=" & "'" & wit_cat(y) & "'"

               bud = dB.Execute(UCase(SQLString))("ttl_bud")

               If Not IsNull(bud) Then bud2 = Val(bud) Else bud2 = 0

               MSHFlexGrid1.TextMatrix(c_row, 3) = Format(bud2, "####.#0")
               MSHFlexGrid1.TextMatrix(c_row, 4) = Format(Val(bud2) - Val(Split(tmp_wit, ",")(1)), "####.#0")

               MSHFlexGrid1.Col = 4
               MSHFlexGrid1.Row = c_row

               If Val(MSHFlexGrid1.TextMatrix(c_row, 4)) > 0 Then
                  MSHFlexGrid1.CellForeColor = vbRed
                Else
                  MSHFlexGrid1.CellForeColor = vbGreen
               End If

               col3_tot = col3_tot + Val(bud2)
               col4_tot = col4_tot + Val(bud2 - Val(Split(tmp_wit, ",")(1)))
               bud2 = 0
            End If

         End If

      Next y
   Next x

   Me.Show
   Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Payments = Nothing

End Sub

Private Sub McToolBar1_Click(ByVal ButtonIndex As Long)

   Select Case ButtonIndex
    Case 30
      Me.Hide
      Unload Me
   End Select

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

