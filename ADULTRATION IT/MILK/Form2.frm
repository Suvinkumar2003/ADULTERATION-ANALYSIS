VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   5895
      Left            =   3000
      OleObjectBlob   =   "Form2.frx":0014
      TabIndex        =   1
      Top             =   840
      Width           =   5055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Admin\Desktop\MEMS BASED ROBOT\memsdata.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MEMS TABLE"
      Top             =   6720
      Width           =   5055
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   360
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GRAPH VIEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'Form1.Show
'Me.Hide
'
'End Sub
'
'
'Private Sub Timer1_Timer()
'
'Data1.Recordset.AddNew
'Data1.Recordset.Fields(0) = Form1.Text2.Text
'Data1.Recordset.Fields(1) = Form1.Text3.Text
'Data1.Recordset.Fields(2) = Form1.Text1.Text
'Data1.Recordset.Update
'Data1.Refresh
''Data1.Recordset.MoveLast
'
'End Sub
