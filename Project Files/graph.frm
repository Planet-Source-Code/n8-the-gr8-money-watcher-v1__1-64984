VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Graph 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3435
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6059
      _Version        =   393216
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483639
      GridColor       =   -2147483634
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
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
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7815
      TabIndex        =   1
      Top             =   60
      Width           =   7815
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
      ButtonIcon20    =   "graph.frx":0000
      ButtonToolTipIcon20=   1
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4215
      Left            =   2340
      OleObjectBlob   =   "graph.frx":0352
      TabIndex        =   2
      Top             =   -150
      Width           =   6700
   End
End
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pass As Integer

Private Sub Form_Load()

  Dim tmp
  Dim a
  Dim ttl As Variant
  Dim x
  Dim i
  Dim xred
  Dim xblue
  Dim xgreen As Integer

   MSChart1.Backdrop.Fill.Brush.FillColor.Set 212, 208, 200

   MSHFlexGrid1.Rows = wit_num
   MSHFlexGrid1.ColWidth(0) = 250
   MSHFlexGrid1.ColWidth(1) = 550
   MSHFlexGrid1.ColWidth(2) = 2300
   MSHFlexGrid1.ColWidth(3) = 300
   MSHFlexGrid1.ColWidth(4) = 300
   MSHFlexGrid1.ColWidth(5) = 300

   For i = 0 To (wit_num - 1)
      MSHFlexGrid1.TextMatrix(i, 2) = UCase((wit_cat(i)))
      MSHFlexGrid1.Row = i
      MSHFlexGrid1.Col = 0
      MSHFlexGrid1.CellFontName = "Marlett"
      MSHFlexGrid1.CellFontSize = 10
      MSHFlexGrid1.Text = "a"
   Next i

   MSHFlexGrid1.Col = 2
   MSHFlexGrid1.Sort = flexSortGenericAscending
   MSChart1.Refresh
   MSChart1.ColumnCount = 1

   a = dB.Execute("SELECT SUM(Amount) as tmp_val FROM MAIN where TRANS LIKE 'WITHDRAWAL%'")("tmp_val")

   If Not IsNull(a) Then
      ttl = a
   End If

   For x = (MSHFlexGrid1.Rows - 1) To 0 Step -1
   
      tmp = dB.Execute("SELECT SUM(Amount) as tmp_val FROM MAIN where Category ='" & MSHFlexGrid1.TextMatrix(x, 2) & "'" & _
      " AND TRANS LIKE 'WITHDRAWAL%'")("tmp_val")
         
      If IsNull(tmp) Then
         MSHFlexGrid1.RemoveItem (x)
      End If

   Next x

   For x = 0 To (MSHFlexGrid1.Rows - 1)

      tmp = dB.Execute("SELECT SUM(Amount) as tmp_val FROM MAIN where Category ='" & MSHFlexGrid1.TextMatrix(x, 2) & "'" & _
         " AND TRANS LIKE 'WITHDRAWAL%'")("tmp_val")

      pass = pass + 1

      If pass > 1 Then
         MSChart1.ColumnCount = MSChart1.ColumnCount + 1
         DoEvents
      End If

      MSChart1.Column = pass
      MSHFlexGrid1.TextMatrix(x, 1) = Format((Val(tmp) / Val(ttl)) * 100, "##.0") & "%"
      MSChart1.Data = (Val(tmp) / Val(ttl)) * 100

      xred = MSChart1.Plot.SeriesCollection.Item(pass).Pen.VtColor.Red
      xblue = MSChart1.Plot.SeriesCollection.Item(pass).Pen.VtColor.Blue
      xgreen = MSChart1.Plot.SeriesCollection.Item(pass).Pen.VtColor.Green

      MSHFlexGrid1.Row = x
      MSHFlexGrid1.Col = 2
      MSHFlexGrid1.CellBackColor = RGB(xred, xgreen, xblue)

      MSHFlexGrid1.TextMatrix(x, 3) = xred
      MSHFlexGrid1.TextMatrix(x, 4) = xgreen
      MSHFlexGrid1.TextMatrix(x, 5) = xblue

   Next x

   Me.Show
   Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Graph = Nothing

End Sub

Private Sub McToolBar1_Click(ByVal ButtonIndex As Long)

   Select Case ButtonIndex

    Case 20
      Me.Hide
      Unload Me
   End Select

End Sub

Public Sub MSChart1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Public Sub MSChart1_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)

End Sub

Private Sub MSHFlexGrid1_Click()

   On Error GoTo ohno
  Dim w
  Dim x
  Dim y
  Dim z
  Dim tmp As String

   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 0) = "a" Then
      MSChart1.SetFocus
      DoEvents

      For x = 0 To MSHFlexGrid1.Rows

         tmp = Replace(MSHFlexGrid1.TextMatrix(x, 1), "%", "")
         w = Val(tmp)

         If MSHFlexGrid1.TextMatrix(x, 0) = "a" And w > 0 Then
            z = z + 1

            If z = (MSHFlexGrid1.RowSel + 1 - y) Then
               MSChart1.SelectPart 7, z, 0, 0, 0
               Exit For
            End If

          Else
            y = y + 1
         End If

      Next x

      Exit Sub
   End If

ohno:

End Sub

Private Sub MSHFlexGrid1_DblClick()

  Dim xred
  Dim xblue
  Dim xgreen As Integer
  Dim tmp
  Dim a
  Dim ttl As Variant
  Dim x

   ttl = 0
   pass = 0
   MSChart1.Refresh
   MSChart1.ColumnCount = 1

   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 0) = "a" Then
      MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 0) = ""
      MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 1) = ""
    Else
      MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 0) = "a"
   End If

   For x = 0 To MSHFlexGrid1.Rows - 1

      If MSHFlexGrid1.TextMatrix(x, 0) = "a" Then
         a = dB.Execute("SELECT SUM(Amount) as tmp_val FROM MAIN where Category ='" & MSHFlexGrid1.TextMatrix(x, 2) & _
            "' AND TRANS LIKE 'Withdrawal%'")("tmp_val")

         If Not IsNull(a) Then
            ttl = ttl + a
         End If

      End If
   Next x

   For x = 0 To MSHFlexGrid1.Rows - 1

      If MSHFlexGrid1.TextMatrix(x, 0) = "a" Then
         tmp = dB.Execute("SELECT SUM(Amount) as tmp_val FROM MAIN where Category ='" & MSHFlexGrid1.TextMatrix(x, 2) & _
            "' AND TRANS LIKE 'Withdrawal%'")("tmp_val")

         If Not IsNull(tmp) Then
            pass = pass + 1

            If pass > 1 Then
               MSChart1.ColumnCount = MSChart1.ColumnCount + 1
            End If

            MSChart1.Column = pass
            MSHFlexGrid1.TextMatrix(x, 1) = Format((Val(tmp) / Val(ttl)) * 100, "##.0") & "%"

            xred = MSHFlexGrid1.TextMatrix(x, 3)
            xgreen = MSHFlexGrid1.TextMatrix(x, 4)
            xblue = MSHFlexGrid1.TextMatrix(x, 5)

            MSChart1.Plot.SeriesCollection(pass).DataPoints(-1).Brush.FillColor.Set xred, xgreen, xblue
            MSChart1.Data = (Val(tmp) / Val(ttl)) * 100

         End If
      End If
   Next x

   DoEvents

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

