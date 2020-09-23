VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Payments 
   BackColor       =   &H80000008&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1740
      ScaleHeight     =   375
      ScaleWidth      =   1455
      TabIndex        =   4
      Top             =   2580
      Width           =   1455
      Begin VB.TextBox txt_month 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
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
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   60
         Width           =   1335
      End
   End
   Begin Money_Watcher.McToolBar McToolBar2 
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   2610
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   529
      BackColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Button_Count    =   5
      ButtonsHeight   =   20
      ButtonsPerRow   =   5
      HoverColor      =   -2147483640
      ToolTipBackCol  =   -2147483640
      BackGradientCol =   -2147483640
      ButtonsMode     =   5
      ShowSeperator   =   0   'False
      ButtonsBackColor=   -2147483640
      ButtonsGradientCol=   -2147483640
      ButtonsGradient =   3
      ButtonCaption1  =   ""
      ButtonIcon1     =   "payments.frx":0000
      ButtonToolTipIcon2=   1
      ButtonToolTipIcon3=   1
      ButtonToolTipIcon4=   1
      ButtonCaption5  =   ""
      ButtonIcon5     =   "payments.frx":0352
      ButtonToolTipIcon5=   1
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
      ButtonIcon30    =   "payments.frx":06A4
      ButtonToolTipIcon30=   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2160
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   3810
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483645
      Cols            =   4
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
      _Band(0).Cols   =   4
   End
End
Attribute VB_Name = "Payments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private current_month As Integer
Private current_year As Integer
Private num As Integer
Private pay As Variant
Private sql_payments As String

Private Sub Form_Load()

  Dim x
  Dim y
  Dim z
  Dim start_date As String
  Dim end_date As String
  Dim found As Boolean

   MSHFlexGrid1.ColWidth(0) = 2750
   MSHFlexGrid1.ColWidth(1) = 900
   MSHFlexGrid1.ColWidth(2) = 840
   MSHFlexGrid1.Cols = 3

   num = CharCount(payment, ",")

   If num = 0 And payment <> "" Then
      num = 1
    ElseIf num > 0 Then
      num = num + 1
   End If

   pay = Split(payment, ",", num)

   For x = 0 To num - 1
      sql_payments = sql_payments & "'" & pay(x) & "'" & ","
   Next x

   sql_payments = UCase(Left(sql_payments, Len(sql_payments) - 1))

   current_month = Month(Date)
   current_year = Year(Date)

   start_date = Month(Date) & "/01/" & Year(Date)
   end_date = Month(Date) & "/" & SetDates(Date) & "/" & Year(Date)
   txt_month.Text = Format(Date, "MMM") & "/" & Format(start_date, "YYYY")

   SQLString = "SELECT CATEGORY, ENTRY_DATE, AMOUNT FROM MAIN" & vbCrLf
   SQLString = SQLString & "WHERE ENTRY_DATE>#" & start_date & "# AND ENTRY_DATE<#" & end_date & "# "
   SQLString = SQLString & "AND CATEGORY IN (" & sql_payments & ") "
   SQLString = SQLString & "AND TRANS='WITHDRAWAL FROM CHECKINGS' "

   Set rs = New ADODB.Recordset
   Set rs = dB.Execute(SQLString)
   Set MSHFlexGrid1.DataSource = rs
   rs.Close
   Set rs = Nothing
   DoEvents

   For x = 0 To num - 1

      found = False

      For y = 1 To MSHFlexGrid1.Rows - 1

         If UCase(MSHFlexGrid1.TextMatrix(y, 0)) = UCase(pay(x)) Then
            found = True
            Exit For
         End If

      Next y

      If found = False Then
         z = MSHFlexGrid1.Rows
         MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
         MSHFlexGrid1.TextMatrix(z, 0) = UCase(pay(x))
      End If

   Next x

   For x = 1 To MSHFlexGrid1.Rows - 1
      Payments.MSHFlexGrid1.TextMatrix(x, 2) = Format(Payments.MSHFlexGrid1.TextMatrix(x, 2), "####.#0")
      Payments.MSHFlexGrid1.TextMatrix(x, 1) = Format(Payments.MSHFlexGrid1.TextMatrix(x, 1), "mm/dd/yyyy")
   Next x

   MSHFlexGrid1.Col = 0
   MSHFlexGrid1.Sort = flexSortGenericAscending
   MSHFlexGrid1.TextMatrix(0, 0) = "Payment"
   MSHFlexGrid1.TextMatrix(0, 1) = "Date"
   MSHFlexGrid1.TextMatrix(0, 2) = "Amount"

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

Private Sub McToolBar2_Click(ByVal ButtonIndex As Long)

  Dim x
  Dim y
  Dim z As Integer
  Dim start_date
  Dim end_date As String
  Dim found As Boolean

   Select Case ButtonIndex

    Case 1
      current_month = current_month - 1

      If current_month = 0 Then
         current_month = 12
         current_year = current_year - 1
      End If

    Case 5
      current_month = current_month + 1

      If current_month = 13 Then
         current_month = 1
         current_year = current_year + 1
      End If

   End Select

   start_date = current_month & "/01/" & current_year
   end_date = current_month & "/" & SetDates(start_date) & "/" & current_year
   txt_month.Text = Format(start_date, "MMM") & "/" & Format(start_date, "YYYY")

   SQLString = "SELECT CATEGORY, ENTRY_DATE, AMOUNT FROM MAIN" & vbCrLf
   SQLString = SQLString & "WHERE ENTRY_DATE>=#" & start_date & "# AND ENTRY_DATE<=#" & end_date & "# "
   SQLString = SQLString & "AND CATEGORY IN (" & sql_payments & ") "
   SQLString = SQLString & "AND TRANS='WITHDRAWAL FROM CHECKINGS' "

   MSHFlexGrid1.Refresh
   Set rs = New ADODB.Recordset
   Set rs = dB.Execute(SQLString)
   Set MSHFlexGrid1.DataSource = rs
   rs.Close
   Set rs = Nothing
   DoEvents

   For x = 0 To num - 1

      found = False

      For y = 1 To MSHFlexGrid1.Rows - 1

         If UCase(MSHFlexGrid1.TextMatrix(y, 0)) = UCase(pay(x)) Then
            found = True
            Exit For
         End If

      Next y

      If found = False Then
         z = MSHFlexGrid1.Rows
         MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
         MSHFlexGrid1.TextMatrix(z, 0) = UCase(pay(x))
      End If

   Next x

   For x = 1 To MSHFlexGrid1.Rows - 1
      Payments.MSHFlexGrid1.TextMatrix(x, 2) = Format(Payments.MSHFlexGrid1.TextMatrix(x, 2), "####.#0")
      Payments.MSHFlexGrid1.TextMatrix(x, 1) = Format(Payments.MSHFlexGrid1.TextMatrix(x, 1), "mm/dd/yyyy")
   Next x

   MSHFlexGrid1.Col = 0
   MSHFlexGrid1.Sort = flexSortGenericAscending
   MSHFlexGrid1.TextMatrix(0, 0) = "Payment"
   MSHFlexGrid1.TextMatrix(0, 1) = "Date"
   MSHFlexGrid1.TextMatrix(0, 2) = "Amount"

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

