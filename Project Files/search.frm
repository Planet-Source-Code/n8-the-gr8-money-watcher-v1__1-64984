VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Search 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   4500
      ScaleHeight     =   210
      ScaleWidth      =   600
      TabIndex        =   30
      Top             =   1660
      Width           =   600
      Begin VB.ComboBox cmb_cond 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         ItemData        =   "search.frx":0000
         Left            =   -30
         List            =   "search.frx":000A
         TabIndex        =   31
         Text            =   "AND"
         Top             =   -30
         Width           =   620
      End
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   4500
      ScaleHeight     =   210
      ScaleWidth      =   600
      TabIndex        =   28
      Top             =   1420
      Width           =   600
      Begin VB.ComboBox cmb_cond 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         ItemData        =   "search.frx":0017
         Left            =   -30
         List            =   "search.frx":0021
         TabIndex        =   29
         Text            =   "AND"
         Top             =   -30
         Width           =   620
      End
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   4500
      ScaleHeight     =   210
      ScaleWidth      =   600
      TabIndex        =   26
      Top             =   1180
      Width           =   600
      Begin VB.ComboBox cmb_cond 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         ItemData        =   "search.frx":002E
         Left            =   -30
         List            =   "search.frx":0038
         TabIndex        =   27
         Text            =   "AND"
         Top             =   -30
         Width           =   620
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   4500
      ScaleHeight     =   210
      ScaleWidth      =   600
      TabIndex        =   24
      Top             =   940
      Width           =   600
      Begin VB.ComboBox cmb_cond 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         ItemData        =   "search.frx":0045
         Left            =   -30
         List            =   "search.frx":004F
         TabIndex        =   25
         Text            =   "AND"
         Top             =   -30
         Width           =   620
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1580
      ScaleHeight     =   210
      ScaleWidth      =   2895
      TabIndex        =   22
      Top             =   1900
      Width           =   2895
      Begin VB.ComboBox search_posted 
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
         ItemData        =   "search.frx":005C
         Left            =   -80
         List            =   "search.frx":0069
         TabIndex        =   23
         Top             =   -30
         Width           =   3010
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   1020
      ScaleHeight     =   210
      ScaleWidth      =   540
      TabIndex        =   20
      Top             =   1900
      Width           =   540
      Begin VB.ComboBox cmb_oper 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         ItemData        =   "search.frx":007C
         Left            =   -60
         List            =   "search.frx":0089
         TabIndex        =   21
         Text            =   "="
         Top             =   -30
         Width           =   590
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   1020
      ScaleHeight     =   210
      ScaleWidth      =   540
      TabIndex        =   18
      Top             =   1660
      Width           =   540
      Begin VB.ComboBox cmb_oper 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         ItemData        =   "search.frx":009A
         Left            =   -60
         List            =   "search.frx":00A7
         TabIndex        =   19
         Text            =   "="
         Top             =   -30
         Width           =   590
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   1020
      ScaleHeight     =   210
      ScaleWidth      =   540
      TabIndex        =   16
      Top             =   1420
      Width           =   540
      Begin VB.ComboBox cmb_oper 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         ItemData        =   "search.frx":00B8
         Left            =   -60
         List            =   "search.frx":00C5
         TabIndex        =   17
         Text            =   "="
         Top             =   -30
         Width           =   590
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   1020
      ScaleHeight     =   210
      ScaleWidth      =   540
      TabIndex        =   14
      Top             =   1180
      Width           =   540
      Begin VB.ComboBox cmb_oper 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         ItemData        =   "search.frx":00D6
         Left            =   -60
         List            =   "search.frx":00EC
         TabIndex        =   15
         Text            =   "="
         Top             =   -30
         Width           =   590
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   1020
      ScaleHeight     =   210
      ScaleWidth      =   540
      TabIndex        =   12
      Top             =   940
      Width           =   540
      Begin VB.ComboBox cmb_oper 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         ItemData        =   "search.frx":0105
         Left            =   -60
         List            =   "search.frx":0112
         TabIndex        =   13
         Text            =   "="
         Top             =   -30
         Width           =   590
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   4500
      ScaleHeight     =   210
      ScaleWidth      =   585
      TabIndex        =   10
      Top             =   700
      Width           =   590
      Begin VB.ComboBox cmb_cond 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         ItemData        =   "search.frx":0123
         Left            =   -30
         List            =   "search.frx":012D
         TabIndex        =   11
         Text            =   "AND"
         Top             =   -30
         Width           =   620
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   1020
      ScaleHeight     =   210
      ScaleWidth      =   540
      TabIndex        =   8
      Top             =   700
      Width           =   540
      Begin VB.ComboBox cmb_oper 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         ItemData        =   "search.frx":013A
         Left            =   -60
         List            =   "search.frx":0150
         TabIndex        =   9
         Text            =   "="
         Top             =   -30
         Width           =   590
      End
   End
   Begin VB.TextBox search_category 
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
      ForeColor       =   &H80000003&
      Height          =   210
      Left            =   1580
      TabIndex        =   3
      Top             =   1430
      Width           =   2895
   End
   Begin VB.TextBox search_trans 
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
      ForeColor       =   &H80000003&
      Height          =   210
      Left            =   1580
      TabIndex        =   1
      Top             =   950
      Width           =   2895
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      ButtonIcon1     =   "search.frx":0169
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
      ButtonIcon30    =   "search.frx":04BB
      ButtonToolTipIcon30=   1
   End
   Begin VB.TextBox search_date 
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
      ForeColor       =   &H80000003&
      Height          =   210
      Left            =   1580
      TabIndex        =   0
      Top             =   700
      Width           =   2895
   End
   Begin VB.TextBox search_amount 
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
      ForeColor       =   &H80000003&
      Height          =   210
      Left            =   1580
      TabIndex        =   2
      Top             =   1190
      Width           =   2895
   End
   Begin VB.TextBox search_description 
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
      ForeColor       =   &H80000003&
      Height          =   210
      Left            =   1580
      TabIndex        =   4
      Top             =   1660
      Width           =   2895
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1740
      Left            =   0
      TabIndex        =   5
      Top             =   420
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   3069
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
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  Dim i As Integer

   Me.Show

   MSHFlexGrid1.ColWidth(0) = 980
   MSHFlexGrid1.ColWidth(1) = 520
   MSHFlexGrid1.ColWidth(2) = 2950
   MSHFlexGrid1.ColWidth(3) = 580
   MSHFlexGrid1.Rows = Header.MSHFlexGrid1.Cols
   MSHFlexGrid1.Cols = 4
   MSHFlexGrid1.TextMatrix(0, 0) = "Field"
   MSHFlexGrid1.TextMatrix(0, 1) = "Oper"
   MSHFlexGrid1.TextMatrix(0, 2) = "Value"
   MSHFlexGrid1.TextMatrix(0, 3) = "Cond"

   For i = 1 To (Header.MSHFlexGrid1.Cols - 1)
      MSHFlexGrid1.TextMatrix(i, 0) = Header.MSHFlexGrid1.ColHeaderCaption(0, i)
   Next i

   MSHFlexGrid1.Rows = Header.MSHFlexGrid1.Cols - 1

   Search.MSHFlexGrid1.TextMatrix(1, 0) = "Entry Date"
   Search.MSHFlexGrid1.TextMatrix(2, 0) = "Transaction"
   Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Search = Nothing

End Sub

Private Sub McToolBar1_Click(ByVal ButtonIndex As Long)

  Dim andor_field As String
  Dim like_field As String

   Select Case ButtonIndex

    Case 1

      SQLString = "SELECT * FROM MAIN WHERE"

      If search_date.Text <> "" Then
         SQLString = SQLString & " ENTRY_DATE" & cmb_oper(0).Text & "#" & Format(search_date.Text, "MM/DD/YY") & "#"
         andor_field = " " & cmb_cond(0).Text
      End If

      If search_trans.Text <> "" Then
         If UCase(cmb_oper(1).Text) = "LIKE" Then
            like_field = "%"
            cmb_oper(1).Text = " like"
         End If

         SQLString = SQLString & andor_field & " TRANS" & cmb_oper(1).Text & "'" & like_field & search_trans.Text & _
               like_field & "'"
         andor_field = " " & cmb_cond(1).Text
         like_field = ""
         cmb_oper(1).Text = Trim(cmb_oper(1).Text)
      End If

      If search_amount.Text <> "" Then
         SQLString = SQLString & andor_field & " AMOUNT" & cmb_oper(2).Text & search_amount.Text
         andor_field = " " & cmb_cond(2).Text
      End If

      If search_category.Text <> "" Then
         If UCase(cmb_oper(3).Text) = "LIKE" Then
            like_field = "%"
            cmb_oper(3).Text = " like"
         End If

         SQLString = SQLString & andor_field & " CATEGORY" & cmb_oper(3).Text & "'" & like_field & search_category.Text _
               & like_field & "'"
         andor_field = " " & cmb_cond(3).Text
         like_field = ""
         cmb_oper(3).Text = Trim(cmb_oper(3).Text)
      End If

      If search_description.Text <> "" Then
         If UCase(cmb_oper(4).Text) = "LIKE" Then
            like_field = "%"
            cmb_oper(4).Text = " like"
         End If

         SQLString = SQLString & andor_field & " DESCRIPTION" & cmb_oper(4).Text & "'" & like_field & _
               search_description.Text & like_field & "'"
         andor_field = " " & cmb_cond(4).Text
         like_field = ""
         cmb_oper(4).Text = Trim(cmb_oper(4).Text)
      End If

      If search_posted.Text <> "" Then
         SQLString = SQLString & andor_field & " POSTED" & cmb_oper(5).Text & search_posted.Text
      End If

      Header.SQL_txt.Text = UCase(SQLString)
      Set rs = New ADODB.Recordset
      Set rs = dB.Execute(UCase(SQLString & " ORDER BY " & Header.default_sort))
      Set Header.MSHFlexGrid1.DataSource = rs
      rs.Close
      Set rs = Nothing

      Call Header.post_load

      Me.Hide
      Unload Me

    Case 30

      Me.Hide
      Unload Me

   End Select

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

