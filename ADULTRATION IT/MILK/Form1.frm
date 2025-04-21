VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   3600
      Top             =   3840
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   360
      Top             =   2760
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   11
      Top             =   7560
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   10
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      TabIndex        =   9
      Top             =   7560
      Width           =   735
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1920
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   840
      Top             =   1800
   End
   Begin VB.Label Label15 
      Caption         =   "M E L A M I N E"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   12000
      TabIndex        =   17
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "C O L O R"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6480
      TabIndex        =   16
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "D E N S I T Y"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1200
      TabIndex        =   15
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "DENSITY"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "COLOR"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "MELAMINE"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   12
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "VARIABLE"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5760
      TabIndex        =   8
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "VARIABLE"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11400
      TabIndex        =   7
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label7 
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   5
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "VARIABLE"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   16080
      TabIndex        =   2
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "         MILK ADULTERATION MONITORING SYSTEM"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   7470
   End
   Begin VB.Line Line7 
      X1              =   15720
      X2              =   15720
      Y1              =   360
      Y2              =   1320
   End
   Begin VB.Line Line47 
      X1              =   14160
      X2              =   17520
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line45 
      X1              =   17520
      X2              =   17520
      Y1              =   360
      Y2              =   8520
   End
   Begin VB.Line Line46 
      X1              =   14160
      X2              =   17520
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line44 
      BorderWidth     =   2
      X1              =   11640
      X2              =   11760
      Y1              =   3720
      Y2              =   3840
   End
   Begin VB.Line Line43 
      BorderWidth     =   2
      X1              =   11640
      X2              =   11520
      Y1              =   3720
      Y2              =   3840
   End
   Begin VB.Line Line42 
      BorderWidth     =   2
      X1              =   11640
      X2              =   11640
      Y1              =   3720
      Y2              =   4440
   End
   Begin VB.Line Line41 
      BorderWidth     =   2
      X1              =   12480
      X2              =   12600
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line Line40 
      BorderWidth     =   2
      X1              =   12480
      X2              =   12360
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line Line39 
      BorderWidth     =   2
      X1              =   16320
      X2              =   16200
      Y1              =   6240
      Y2              =   6360
   End
   Begin VB.Line Line38 
      BorderWidth     =   2
      X1              =   16320
      X2              =   16200
      Y1              =   6240
      Y2              =   6120
   End
   Begin VB.Line Line37 
      BorderWidth     =   2
      X1              =   15960
      X2              =   15840
      Y1              =   6480
      Y2              =   6600
   End
   Begin VB.Line Line36 
      BorderWidth     =   2
      X1              =   15960
      X2              =   15840
      Y1              =   6480
      Y2              =   6360
   End
   Begin VB.Line Line35 
      BorderWidth     =   2
      X1              =   15240
      X2              =   15960
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line34 
      BorderWidth     =   2
      X1              =   12480
      X2              =   16320
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line33 
      BorderWidth     =   2
      X1              =   12480
      X2              =   12480
      Y1              =   3120
      Y2              =   6240
   End
   Begin VB.Line Line32 
      X1              =   14280
      X2              =   17520
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line31 
      BorderWidth     =   2
      X1              =   10680
      X2              =   10560
      Y1              =   6240
      Y2              =   6360
   End
   Begin VB.Line Line30 
      BorderWidth     =   2
      X1              =   10680
      X2              =   10560
      Y1              =   6240
      Y2              =   6120
   End
   Begin VB.Line Line29 
      BorderWidth     =   2
      X1              =   10320
      X2              =   10200
      Y1              =   6480
      Y2              =   6600
   End
   Begin VB.Line Line28 
      BorderWidth     =   2
      X1              =   10320
      X2              =   10200
      Y1              =   6480
      Y2              =   6360
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4800
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line27 
      BorderWidth     =   2
      X1              =   9600
      X2              =   10320
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   960
      X2              =   960
      Y1              =   3480
      Y2              =   4080
   End
   Begin VB.Line Line26 
      BorderWidth     =   2
      X1              =   6240
      X2              =   6360
      Y1              =   3720
      Y2              =   3840
   End
   Begin VB.Line Line25 
      BorderWidth     =   2
      X1              =   6240
      X2              =   6120
      Y1              =   3720
      Y2              =   3840
   End
   Begin VB.Line Line24 
      BorderWidth     =   2
      X1              =   6240
      X2              =   6240
      Y1              =   3720
      Y2              =   4320
   End
   Begin VB.Line Line23 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6960
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line Line22 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6720
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   6840
      X2              =   10680
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   3120
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   4800
      X2              =   4680
      Y1              =   6480
      Y2              =   6600
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4800
      X2              =   4680
      Y1              =   6480
      Y2              =   6360
   End
   Begin VB.Line Line19 
      X1              =   5280
      X2              =   5160
      Y1              =   6240
      Y2              =   6360
   End
   Begin VB.Line Line18 
      X1              =   5280
      X2              =   5160
      Y1              =   6240
      Y2              =   6120
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   1560
      X2              =   1680
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   1560
      X2              =   1440
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   960
      X2              =   1080
      Y1              =   3480
      Y2              =   3600
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   960
      X2              =   840
      Y1              =   3480
      Y2              =   3600
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   1560
      X2              =   5280
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   3120
      Y2              =   6240
   End
   Begin VB.Line Line6 
      X1              =   1920
      X2              =   1920
      Y1              =   360
      Y2              =   1320
   End
   Begin VB.Line Line5 
      X1              =   360
      X2              =   14160
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   14280
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   14280
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   360
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   600
      Y2              =   8520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim vi, sival As Double, i As Integer
'Dim imgcapt As Boolean, sdate As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim Val1 As Integer, Val2 As Integer, Val3 As Integer, Val4 As Integer, VAL0, VAL5, VAL6 As Integer
Dim buf As String, Out As Integer, OUT1 As Integer, A As Integer
Dim B As Integer, C As Integer, PORTD As Integer, DOUT As String
Dim SX As Integer, SX1 As Integer, SX2 As Integer
Dim SY As Integer, SY1 As Integer, SY2 As Integer
Dim EX As Integer, EX1 As Integer, EX2 As Integer
Dim EY As Integer, EY1 As Integer, EY2 As Integer
Option Explicit



Private Sub Command1_Click()
Me.Hide
Form2.Show
End Sub

Private Sub Form_Load()

MSComm1.PortOpen = True
MSComm1.Output = "{24}"
    Sleep 100
MSComm1.Output = "{1C80}"
Sleep (100)
MSComm1.Output = "{1D00}"
Sleep (100)
MSComm1.Output = "{1AFF}"
Sleep (100)
SX = Line9.X1
    SY = Line9.Y1
    
     SX1 = Line21.X1
    SY1 = Line21.Y1
    
     SX2 = Line34.X1
    SY2 = Line34.Y1
    
    'MSComm1.Output = "{5DFF}"

'Form2.Data1.DatabaseName = App.Path & "\memsdata.mdb"
'Form2.Data1.RecordSource = "MEMS TABLE"
End Sub
Private Sub Form_Unload(Cancel As Integer)
MSComm1.Output = "{5D00}"
Sleep (1000)
End
End Sub



Private Sub Timer1_Timer()
    Label2.Caption = Date
Label3.Caption = Time
   Val1 = Analog(4) - 900
    Val2 = Analog(5) - 900
    Val3 = Analog(5) - 900
    Val1 = Val1
    If Val1 < 0 Then
    Val1 = 0
    End If

Val1 = Val1

Text1.Text = Val1

Text2.Text = Val2
Text3.Text = Val3
    
 
  
'GRAPH 1
EX = SX + 100
EY = Line8.Y2 - (Val2 / 150) * (Line8.Y2 - Line8.Y1)

Line (SX, SY)-(EX, EY), vbRed
SX = EX
SY = EY


If (SX > Line9.X2 - 50) Then
  Line (Line8.X1, Line8.Y1)-(Line9.X2, Line8.Y2), Me.BackColor, BF
  SX = Line9.X1
  SY = Line9.Y1
  Line8.Refresh
  Line9.Refresh
  Line17.Refresh
  Line18.Refresh
  End If
  
  ' GRAPH 2
  EX1 = SX1 + 100
  EY1 = Line20.Y2 - (Val3 / 150) * (Line20.Y2 - Line20.Y1)

Line (SX1, SY1)-(EX1, EY1), vbRed
SX1 = EX1
SY1 = EY1


If (SX1 > Line21.X2 - 50) Then
  Line (Line20.X1, Line20.Y1)-(Line21.X2, Line20.Y2), Me.BackColor, BF
  SX1 = Line21.X1
  SY1 = Line21.Y1
  Line20.Refresh
  Line21.Refresh
  Line23.Refresh
  Line30.Refresh
  
End If


' GRAPH 3
  EX2 = SX2 + 100
  EY2 = Line33.Y2 - (Val1 / 150) * (Line33.Y2 - Line33.Y1)

Line (SX2, SY2)-(EX2, EY2), vbRed
SX2 = EX2
SY2 = EY2


If (SX2 > Line34.X2 - 50) Then
  Line (Line33.X1, Line33.Y1)-(Line34.X2, Line33.Y2), Me.BackColor, BF
  SX2 = Line34.X1
  SY2 = Line34.Y1
  Line33.Refresh
  Line34.Refresh
  Line41.Refresh
  Line38.Refresh
  
End If
End Sub
Function Analog(no As Integer)
    MSComm1.Output = "{4" & CStr(no) & "}"
    Sleep 100
    buf = MSComm1.Input
    If (buf <> "") Then
        Analog = CInt(Mid$(buf, 2, 4))
    Else
        Analog = 0
    End If
End Function
Private Sub Timer2_Timer()
If Val(Text1.Text) > 420 Then
 Out = Out Or &H1
 Text1.BackColor = vbGreen
Else
  Out = Out And &HFE
  Text1.BackColor = vbRed
End If
  If Val(Text2.Text) > 420 Then
  Out = Out Or &H2
  Text2.BackColor = vbGreen
Else
    Out = Out And &HFD
    Text2.BackColor = vbRed
    End If
    If Val(Text3.Text) > 350 Then
  Out = Out Or &H4
  Text3.BackColor = vbGreen
Else
    Out = Out And &HFB
    Text3.BackColor = vbRed
    End If
    
    
'     Out = Out Or &H10

'    Out = Out And &HEF

If Len(CStr(Hex(Out))) <> 2 Then
    MSComm1.Output = "{5D0" & CStr(Hex(Out)) & "}"
    Sleep (100)
Else
    MSComm1.Output = "{5D" & CStr(Hex(Out)) & "}"
    Sleep (100)
End If
End Sub


