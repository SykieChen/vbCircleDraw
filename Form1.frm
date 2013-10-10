VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ready"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17115
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   17115
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   4560
      TabIndex        =   6
      Text            =   "500000"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1800
      TabIndex        =   4
      Text            =   "10"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   10080
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      FillStyle       =   0  'Solid
      Height          =   6855
      Left            =   120
      ScaleHeight     =   453
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1125
      TabIndex        =   1
      Top             =   480
      Width           =   16935
   End
   Begin VB.Label Label2 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Dot count:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flag As Boolean

Private Sub Command1_Click()
Randomize
Picture1.Cls
Dim i As Integer
Dim s As String
Dim ii As Long
Dim px() As Integer
Dim py() As Integer
ReDim px(CInt(Text1.Text) - 1)
ReDim py(CInt(Text1.Text) - 1)
For i = 0 To (CInt(Text1.Text) - 1) Step 1
    px(i) = Rnd * Picture1.ScaleWidth
    py(i) = Rnd * Picture1.ScaleHeight
Next i
For i = 0 To (CInt(Text1.Text) - 1) Step 1
    Picture1.Circle (px(i), py(i)), 1
Next i
For i = 0 To (CInt(Text1.Text) - 2) Step 1
    Picture1.Line (px(i), py(i))-(px(i + 1), py(i + 1))
Next i
Picture1.Line (px(0), py(0))-(px(i), py(i))
flag = False
Me.Caption = "Working..."
Dim count As Long
count = 0
Do While flag = False
    count = count + 1
    Command1.Caption = "Loop " & count
    DoEvents
    px0 = px(0)
    py0 = py(0)
    For i = 0 To (CInt(Text1.Text) - 2) Step 1
        px(i) = (px(i) + px(i + 1)) / 2
        py(i) = (py(i) + py(i + 1)) / 2
    Next i
    px(i) = (px(i) + px0) / 2
    py(i) = (py(i) + py0) / 2
    Picture1.Cls
    For i = 0 To (CInt(Text1.Text) - 1) Step 1
        Picture1.Circle (px(i), py(i)), 1
    Next i
    For i = 0 To (CInt(Text1.Text) - 2) Step 1
        Picture1.Line (px(i), py(i))-(px(i + 1), py(i + 1))
    Next i
    Picture1.Line (px(0), py(0))-(px(i), py(i))
    For ii = 0 To CLng(Text2.Text) Step 1
        DoEvents
    Next ii
Loop
Me.Caption = "Ready"
Command1.Caption = "Start"
End Sub

Private Sub Command2_Click()
flag = True
End Sub
