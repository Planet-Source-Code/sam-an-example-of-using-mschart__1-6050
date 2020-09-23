VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   5760
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Top             =   3360
      Width           =   4935
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5640
      TabIndex        =   10
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5640
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5640
      TabIndex        =   8
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5640
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Data"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin MSChartLib.MSChart MSChart1 
      Height          =   3255
      Left            =   0
      OleObjectBlob   =   "Chart.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "1999"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "1998"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "1997"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "1996"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1995"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
    MSChart1.chartType = Combo1.ListIndex
    MSChart1.Refresh
End Sub



Private Sub Command2_Click()
 
MSChart1.TitleText = "My Chart" 'The chart title
MSChart1.Row = 1 'The first row in the chart
MSChart1.Data = Val(Text1) 'Validating the data that will be entered into Text1
MSChart1.RowLabel = "1995" 'The name of the first row

MSChart1.Row = 2
MSChart1.Data = Val(Text2)
MSChart1.RowLabel = "1996"

MSChart1.Row = 3
MSChart1.Data = Val(Text3)
MSChart1.RowLabel = "1997"

MSChart1.Row = 4
MSChart1.Data = Val(Text4)
MSChart1.RowLabel = "1998"

MSChart1.Row = 5
MSChart1.Data = Val(Text5)
MSChart1.RowLabel = "1999"

'If you want to add another bar with the exisitng one
MSChart1.Column = 2
MSChart1.Data = 10
MSChart1.Data = 20
MSChart1.Data = 30
MSChart1.Data = 40
MSChart1.Data = 50
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
'these are for the various types of bars
Combo1.AddItem "3dbar"
Combo1.AddItem "2dbar"
Combo1.AddItem "3dline"
Combo1.AddItem "2dline"
Combo1.AddItem "3darea"
Combo1.AddItem "2dstep"
Combo1.AddItem "3dstep"

End Sub
