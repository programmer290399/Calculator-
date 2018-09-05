VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.1#0"; "vbskpro2.ocx"
Object = "{69C832A0-68F4-452F-9B16-837E157288D9}#1.0#0"; "Styler_button.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALCULATOR "
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin glxpbuttonz.UserButtonz UserButtonz14 
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "."
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin Styler_button.StylerButton StylerButton3 
      Height          =   495
      Left            =   2880
      TabIndex        =   21
      Top             =   5520
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "MAKE INT"
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Styler_button.StylerButton StylerButton2 
      Height          =   495
      Left            =   5160
      TabIndex        =   20
      Top             =   5520
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "SQUARE ROOT"
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Styler_button.StylerButton StylerButton1 
      Height          =   495
      Left            =   600
      TabIndex        =   19
      Top             =   5520
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "ROUND OFF"
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin glxpbuttonz.UserButtonz UserButtonz13 
      Height          =   615
      Left            =   4320
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LESS OPTIONS"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   160
      ColorButtonUp   =   128
      ColorButtonDown =   240
      BorderBrightness=   0
      ColorBright     =   255
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin glxpbuttonz.UserButtonz UserButtonz12 
      Height          =   615
      Left            =   4320
      TabIndex        =   17
      Top             =   4320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "MORE OPTIONS"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   160
      ColorButtonUp   =   128
      ColorButtonDown =   240
      BorderBrightness=   0
      ColorBright     =   255
      DisplayHand     =   -1  'True
      ColorScheme     =   3
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   2760
      Top             =   4200
      _ExtentX        =   1270
      _ExtentY        =   1270
      SysDisableSkinCaption=   "&Disable Skin"
   End
   Begin glxpbuttonz.UserButtonz UserButtonz11 
      Height          =   735
      Left            =   6840
      TabIndex        =   16
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CLEAR "
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   33023
      ColorButtonUp   =   33023
      ColorButtonDown =   33023
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Styler_button.StylerButton cmddivide 
      Height          =   975
      Left            =   4080
      TabIndex        =   15
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
      Caption         =   "/"
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Styler_button.StylerButton StylerButton6 
      Height          =   975
      Left            =   4080
      TabIndex        =   14
      Top             =   3120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
      Caption         =   "="
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Styler_button.StylerButton cmdsubtract 
      Height          =   975
      Left            =   5640
      TabIndex        =   13
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      Caption         =   "-"
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Styler_button.StylerButton cmdmultiply 
      Height          =   975
      Left            =   4080
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      Caption         =   "X"
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Styler_button.StylerButton cmdadd 
      Height          =   3855
      Left            =   7200
      TabIndex        =   11
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   6800
      Caption         =   "+"
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin glxpbuttonz.UserButtonz UserButtonz10 
      Height          =   975
      Left            =   1320
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin glxpbuttonz.UserButtonz UserButtonz9 
      Height          =   975
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "9"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin glxpbuttonz.UserButtonz UserButtonz8 
      Height          =   975
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "8"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin glxpbuttonz.UserButtonz UserButtonz7 
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "7"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin glxpbuttonz.UserButtonz UserButtonz6 
      Height          =   975
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "6"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin glxpbuttonz.UserButtonz UserButtonz5 
      Height          =   975
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "5"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin glxpbuttonz.UserButtonz UserButtonz4 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "4"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin glxpbuttonz.UserButtonz UserButtonz3 
      Height          =   975
      Left            =   2520
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "3"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin glxpbuttonz.UserButtonz UserButtonz2 
      Height          =   975
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin glxpbuttonz.UserButtonz UserButtonz1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OP As String
Dim NUM1 As Double
Dim NUM2 As Double



Private Sub cmdadd_Click()
NUM1 = Text1.Text
Text1.Text = " "
OP = "+"

End Sub

Private Sub cmddivide_Click()
NUM1 = Text1.Text
Text1.Text = ""
OP = "/"



End Sub

Private Sub cmdmultiply_Click()
NUM1 = Text1.Text
Text1.Text = ""
OP = "x"
End Sub

Private Sub cmdsubtract_Click()
NUM1 = Text1.Text
Text1.Text = ""
OP = "-"
End Sub

Private Sub StylerButton1_Click()
A = Val(Text1.Text)
B = CInt(A)
Text1.Text = B



End Sub

Private Sub StylerButton2_Click()
A = Val(Text1.Text)
B = Sqr(A)
Text1.Text = B

End Sub

Private Sub StylerButton3_Click()
A = Val(Text1.Text)
B = Int(A)
Text1.Text = B

End Sub

Private Sub StylerButton6_Click()
NUM2 = Text1.Text
If OP = "+" Then
Text1.Text = NUM1 + NUM2
ElseIf OP = "-" Then
Text1.Text = NUM1 - NUM2
ElseIf OP = "/" Then
Text1.Text = NUM1 / NUM2
ElseIf OP = "x" Then
Text1.Text = NUM1 * NUM2
End If




End Sub

Private Sub UserButtonz1_Click()
Text1.Text = Text1.Text & "1"
End Sub

Private Sub UserButtonz10_Click()
Text1.Text = Text1.Text & "0"
End Sub


Private Sub UserButtonz11_Click()
Text1.Text = ""

End Sub

Private Sub UserButtonz12_Click()
UserButtonz13.Visible = True
UserButtonz12.Visible = False
Do Until Me.Height > 6660
Me.Height = Me.Height + 1
DoEvents
Loop

End Sub

Private Sub UserButtonz13_Click()
Me.Height = 5550
UserButtonz12.Visible = True
UserButtonz13.Visible = False

End Sub

Private Sub UserButtonz14_Click()
Text1.Text = Text1.Text & "."
End Sub

Private Sub UserButtonz2_Click()
Text1.Text = Text1.Text & "2"
End Sub

Private Sub UserButtonz3_Click()

Text1.Text = Text1.Text & "3"
End Sub

Private Sub UserButtonz4_Click()

Text1.Text = Text1.Text & "4"
End Sub

Private Sub UserButtonz5_Click()
Text1.Text = Text1.Text & "5"
End Sub

Private Sub UserButtonz6_Click()
Text1.Text = Text1.Text & "6"
End Sub

Private Sub UserButtonz7_Click()
Text1.Text = Text1.Text & "7"
End Sub

Private Sub UserButtonz8_Click()
Text1.Text = Text1.Text & "8"
End Sub

Private Sub UserButtonz9_Click()
Text1.Text = Text1.Text & "9"
End Sub
