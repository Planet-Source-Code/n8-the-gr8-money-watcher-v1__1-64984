Attribute VB_Name = "Primary"
Option Explicit
Public Declare Function SendMessageX Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" _
      Alias "SendMessageA" ( _
      ByVal hwnd As Long, _
      ByVal wMsg As Long, _
      ByVal wParam As Integer, _
      ByVal iparam As Long) As Long
Declare Sub SetWindowPos Lib "user32" ( _
      ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" _
      Alias "GetPrivateProfileStringA" ( _
      ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, _
      ByVal lpDefault As String, _
      ByVal lpReturnedString As String, _
      ByVal nSize As Long, _
      ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" _
      Alias "WritePrivateProfileStringA" ( _
      ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, _
      ByVal lpString As Any, _
      ByVal lpFileName As String) As Long
Private Declare Function GetWindowLong Lib "user32" _
      Alias "GetWindowLongA" ( _
      ByVal hwnd As Long, _
      ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
      Alias "SetWindowLongA" ( _
      ByVal hwnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) As Long

Public SQLString As String
Public DEFAULT_SQLString As String
Public option_file As String
Public rem_edit_mode As Boolean
Public dB As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rem_dB As New ADODB.Connection
Public rem_dB2 As New ADODB.Connection
Public reminders_file As String
Public tmp_col As Integer
Public withdrawals As String
Public wit_cat As Variant
Public wit_num As Integer
Public deposits As String
Public dep_cat As Variant
Public dep_num As Integer
Public payment As String
Public pay_cat As Variant
Public pay_num As Integer
Public payment_exclude As String
Public no_pay_cat As Variant
Public no_pay_num As Integer
Public budget_timeframes As String
Public budget_num As Integer
Public budget_cat As Variant
Public transfers As String
Public xfer_cat As Variant
Public xfer_num As Integer
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const CB_FINDSTRING = &H14C

Public Type POINTAPI
   x As Long
   y As Long
End Type

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000

Private Const SWP_NOZORDER     As Long = &H4
Private Const SWP_NOACTIVATE   As Long = &H10
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_REFRESH      As Long = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_FRAMECHANGED

Public Function AutoComplete(cbCombo As ComboBox, sKeyAscii As Integer, Optional bUpperCase As Boolean = True) As Integer

  Dim lngFind As Long
  Dim intPos As Integer
  Dim intLength As Integer
  Dim tStr As String

   With cbCombo

      If sKeyAscii = 8 Then
         If .SelStart = 0 Then Exit Function
         .SelStart = .SelStart - 1
         .SelLength = 32000
         .SelText = ""
       Else
         intPos = .SelStart '// save intial cursor position
         tStr = .Text '// save String

         If bUpperCase = True Then
            .SelText = UCase(Chr(sKeyAscii)) '// change string. (uppercase only)
          Else
            .SelText = UCase(Chr(sKeyAscii)) '// change string. (leave case alone)
         End If

      End If

      lngFind = SendMessageX(.hwnd, CB_FINDSTRING, 0, ByVal .Text) '// Find string in combobox

      If lngFind = -1 Then '// if String Not found
         .Text = tStr '// Set old String (used For boxes that require charachter monitoring
         .SelStart = intPos '// Set cursor position
         .SelLength = (Len(.Text) - intPos) '// Set selected length
         AutoComplete = 0 '// return 0 value to KeyAscii
         Exit Function

       Else '// If String found
         intPos = .SelStart '// save cursor position
         intLength = Len(.List(lngFind)) - Len(.Text) '// save remaining highlighted text length
         .SelText = .SelText & Right(.List(lngFind), intLength) '// change new text in String
         '.Text = .List(lngFind)'// Use this inst
         '     ead of the above .Seltext line to set th
         '     e text typed to the exact case of the it
         '     em selected in the combo box.
         .SelStart = intPos '// Set cursor position
         .SelLength = intLength '// Set selected length
      End If

   End With

End Function

Function CharCount(OrigString As String, Chars As String, Optional CaseSensitive As Boolean = False) As Long

  Dim lLen As Long
  Dim lCharLen As Long
  Dim lAns As Long
  Dim sInput As String
  Dim sChar As String
  Dim lCtr As Long
  Dim lEndOfLoop As Long
  Dim bytCompareType As Byte

   sInput = OrigString
   If sInput = "" Then Exit Function
   lLen = Len(sInput)
   lCharLen = Len(Chars)
   lEndOfLoop = (lLen - lCharLen) + 1
   bytCompareType = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)

   For lCtr = 1 To lEndOfLoop
      sChar = Mid(sInput, lCtr, lCharLen)
      If StrComp(sChar, Chars, bytCompareType) = 0 Then lAns = lAns + 1
   Next

   CharCount = lAns

End Function

Public Sub FormMove(theform As Form)

   ReleaseCapture
   Call SendMessage(theform.hwnd, &HA1, 2, 0&)

End Sub

Public Function GET_transactions(opt As String)

   Select Case opt

    Case "1"
      withdrawals = readINI("Defaults", "withdrawals", option_file)
      wit_num = CharCount(withdrawals, ",")

      If wit_num = 0 And withdrawals <> "" Then
         wit_num = 1
       ElseIf wit_num > 0 Then
         wit_num = wit_num + 1
      End If

      wit_cat = Split(withdrawals, ",", wit_num)

    Case "2"

      deposits = readINI("Defaults", "deposits", option_file)
      dep_num = CharCount(deposits, ",")

      If dep_num = 0 And deposits <> "" Then
         dep_num = 1
       ElseIf dep_num > 0 Then
         dep_num = dep_num + 1
      End If

      dep_cat = Split(deposits, ",", dep_num)

    Case "3"
      payment_exclude = readINI("Defaults", "payments_exclude", option_file)
      no_pay_num = CharCount(payment_exclude, ",")

      If no_pay_num = 0 And payment_exclude <> "" Then
         no_pay_num = 1
       ElseIf no_pay_num > 0 Then
         no_pay_num = no_pay_num + 1
      End If

      no_pay_cat = Split(payment_exclude, ",", no_pay_num)

    Case "4"
      payment = readINI("Defaults", "payments", option_file)
      pay_num = CharCount(payment, ",")

      If pay_num = 0 And payment <> "" Then
         pay_num = 1
       ElseIf pay_num > 0 Then
         pay_num = pay_num + 1
      End If

      pay_cat = Split(payment, ",", pay_num)

    Case "5"

      budget_timeframes = readINI("Defaults", "budget_timeframes", option_file)
      budget_num = CharCount(budget_timeframes, ",")

      If budget_num = 0 And budget_timeframes <> "" Then
         budget_num = 1
       ElseIf pay_num > 0 Then
         budget_num = budget_num + 1
      End If

      budget_cat = Split(budget_timeframes, ",", budget_num)

    Case "6"

      transfers = readINI("Defaults", "transfers", option_file)
      xfer_num = CharCount(transfers, ",")

      If xfer_num = 0 And transfers <> "" Then
         xfer_num = 1
       ElseIf xfer_num > 0 Then
         xfer_num = xfer_num + 1
      End If

      xfer_cat = Split(transfers, ",", xfer_num)

    Case "ALL"

      withdrawals = readINI("Defaults", "withdrawals", option_file)
      wit_num = CharCount(withdrawals, ",")

      If wit_num = 0 And withdrawals <> "" Then
         wit_num = 1
       ElseIf wit_num > 0 Then
         wit_num = wit_num + 1
      End If

      wit_cat = Split(withdrawals, ",", wit_num)

      deposits = readINI("Defaults", "deposits", option_file)
      dep_num = CharCount(deposits, ",")

      If dep_num = 0 And deposits <> "" Then
         dep_num = 1
       ElseIf dep_num > 0 Then
         dep_num = dep_num + 1
      End If

      dep_cat = Split(deposits, ",", dep_num)

      payment_exclude = readINI("Defaults", "payments_exclude", option_file)
      no_pay_num = CharCount(payment_exclude, ",")

      If no_pay_num = 0 And payment_exclude <> "" Then
         no_pay_num = 1
       ElseIf no_pay_num > 0 Then
         no_pay_num = no_pay_num + 1
      End If

      no_pay_cat = Split(payment_exclude, ",", no_pay_num)

      payment = readINI("Defaults", "payments", option_file)
      pay_num = CharCount(payment, ",")

      If pay_num = 0 And payment <> "" Then
         pay_num = 1
       ElseIf pay_num > 0 Then
         pay_num = pay_num + 1
      End If

      pay_cat = Split(payment, ",", pay_num)

      budget_timeframes = readINI("Defaults", "budget_timeframes", option_file)
      budget_num = CharCount(budget_timeframes, ",")

      If budget_num = 0 And budget_timeframes <> "" Then
         budget_num = 1
       ElseIf pay_num > 0 Then
         budget_num = budget_num + 1
      End If

      budget_cat = Split(budget_timeframes, ",", budget_num)

      transfers = readINI("Defaults", "transfers", option_file)
      xfer_num = CharCount(transfers, ",")

      If xfer_num = 0 And transfers <> "" Then
         xfer_num = 1
       ElseIf xfer_num > 0 Then
         xfer_num = xfer_num + 1
      End If

      xfer_cat = Split(transfers, ",", xfer_num)

   End Select

End Function

Function readINI(Section, KeyName, filename As String) As String

  Dim sRet As String

   sRet = String(9999, Chr(0))
   readINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))

End Function

Public Function SetDates(tmp_date) As String

   Select Case Month(tmp_date)
    Case 1, 3, 5, 7, 8, 10, 12
      SetDates = "31"

    Case 4, 6, 9, 11
      SetDates = "30"

    Case 2

      If (Year(tmp_date) Mod 4) = 0 Then
         SetDates = "29"
       Else
         SetDates = "28"
      End If

   End Select

End Function

Public Property Let ShowCaption(ByRef Form As Form, ByVal New_Value As Boolean)

  Dim lngStyle As Long

   lngStyle = GetWindowLong(Form.hwnd, GWL_STYLE)

   If (New_Value) Then
      lngStyle = lngStyle Or WS_CAPTION
    Else
      lngStyle = lngStyle And Not WS_CAPTION
   End If

   SetWindowLong Form.hwnd, GWL_STYLE, lngStyle

   SetWindowPos Form.hwnd, 0, 0, 0, 0, 0, SWP_REFRESH

End Property

Public Property Get ShowCaption(ByRef Form As Form) As Boolean

   ShowCaption = (GetWindowLong(Form.hwnd, GWL_STYLE) And WS_CAPTION) = WS_CAPTION

End Property

Public Function writeINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer

  Dim r

   r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)

End Function

