VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Datapad 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox rtb_datapad 
      Height          =   3135
      Left            =   -20
      TabIndex        =   2
      Top             =   420
      Width           =   4715
      _ExtentX        =   8308
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"datapad.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   4215
      TabIndex        =   1
      Top             =   60
      Width           =   4215
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   30
         X2              =   6250
         Y1              =   150
         Y2              =   150
      End
   End
   Begin Money_Watcher.McToolBar McToolBar1 
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
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
      ButtonsSeperatorWidth=   14
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
      ButtonIcon20    =   "datapad.frx":007B
      ButtonToolTipIcon20=   1
   End
End
Attribute VB_Name = "Datapad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private datapad_file As String
Private Sub Form_Load()

   EnableURLDetect rtb_datapad.hwnd, Me.hwnd
   datapad_file = readINI("Defaults", "datapad_file", option_file)

   Open datapad_file For Input As #1
   rtb_datapad = Input(LOF(1), 1)
   Close #1
   DoEvents

   Me.Show
   Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set Datapad = Nothing

End Sub

Private Sub McToolBar1_Click(ByVal ButtonIndex As Long)

   Select Case ButtonIndex

    Case 20
      Open datapad_file For Output As #1
      Print #1, rtb_datapad
      Close #1
      Me.Hide
      Call DisableURLDetect
      Unload Me

   End Select

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   FormMove Me

End Sub

