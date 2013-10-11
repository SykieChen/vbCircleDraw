VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ready"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   17160
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "+"
      Height          =   255
      Left            =   9840
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-"
      Height          =   255
      Left            =   9600
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Help..."
      Height          =   255
      Left            =   15720
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   8880
      TabIndex        =   14
      Text            =   "1"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   5040
      TabIndex        =   10
      Text            =   "1"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   2760
      TabIndex        =   6
      Text            =   "100000"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   1200
      TabIndex        =   4
      Text            =   "10"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   13800
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   11880
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
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
   Begin VB.Label Label6 
      Caption         =   "Loop 0"
      Height          =   255
      Left            =   10200
      TabIndex        =   18
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Draw Frequency:"
      Height          =   255
      Left            =   7440
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Zoom:"
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Dot Size:"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Timeout:"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Dot Count:"
      Height          =   255
      Left            =   240
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
Dim px() As Double
Dim py() As Double
Dim maxx As Double
Dim maxy As Double
Dim minx As Double
Dim miny As Double
Dim zoomx As Double
Dim zoomy As Double
Dim zoom As Double
Dim zoomtotal As Double
zoomtotal = 1
ReDim px(CInt(Text1.Text) - 1)
ReDim py(CInt(Text1.Text) - 1)
For i = 0 To (CInt(Text1.Text) - 1) Step 1
    px(i) = Rnd * Picture1.ScaleWidth
    py(i) = Rnd * Picture1.ScaleHeight
Next i
For i = 0 To (CInt(Text1.Text) - 1) Step 1
    Picture1.Circle (px(i), py(i)), CInt(Text3.Text)
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
    Label6.Caption = "Loop " & count
    DoEvents
    minx = px(0)
    miny = py(0)
    For i = 0 To (CInt(Text1.Text) - 2) Step 1
        px(i) = (px(i) + px(i + 1)) / 2
        py(i) = (py(i) + py(i + 1)) / 2
    Next i
    px(i) = (px(i) + minx) / 2
    py(i) = (py(i) + miny) / 2
    'zoom
    'init
    minx = Picture1.ScaleWidth
    miny = Picture1.ScaleHeight
    maxx = 0
    maxy = 0
    'finger out min x min y
    For i = 0 To (CInt(Text1.Text) - 1) Step 1
        If px(i) <= minx Then minx = px(i)
        If py(i) <= miny Then miny = py(i)
        If px(i) >= maxx Then maxx = px(i)
        If py(i) >= maxy Then maxy = py(i)
    Next i
    zoomx = Picture1.ScaleWidth / (maxx - minx)
    zoomy = Picture1.ScaleHeight / (maxy - miny)
    If zoomx < zoomy Then zoom = zoomx Else zoom = zoomy
    zoomtotal = zoomtotal * zoom
    Rem zoom value will only be refreshed while drawing
    For i = 0 To (CInt(Text1.Text) - 1) Step 1
        px(i) = px(i) - minx
        py(i) = py(i) - miny
        px(i) = px(i) * zoom
        py(i) = py(i) * zoom
    Next i
    'draw
    If (count Mod CInt(Text5.Text) = 0) Then
        Text4.Text = Int(zoomtotal * 10) / 10
        Picture1.Cls
        For i = 0 To (CInt(Text1.Text) - 1) Step 1
            Picture1.Circle (px(i), py(i)), CInt(Text3.Text)
        Next i
        For i = 0 To (CInt(Text1.Text) - 2) Step 1
            Picture1.Line (px(i), py(i))-(px(i + 1), py(i + 1))
        Next i
        Picture1.Line (px(0), py(0))-(px(i), py(i))
    End If
    'timeout
    For ii = 0 To CLng(Text2.Text) Step 1
        DoEvents
    Next ii
Loop
Me.Caption = "Ready"
End Sub

Private Sub Command2_Click()
flag = True
End Sub

Private Sub Command3_Click()
If (CLng(Text2.Text) / 10) >= 1 Then
    Text2.Text = CLng(Text2.Text) / 10
Else
    Text2.Text = "1"
End If
End Sub

Private Sub Command4_Click()
Text2.Text = CLng(Text2.Text) * 10
End Sub


Private Sub Command5_Click()
MsgBox "使用说明：" & Chr(10) & Chr(10) & "Timeout：运算间隔。计算每组坐标之间的超时，为了方便观察可以调大，最小为1" & Chr(10) & "Draw Frequency：绘图频率。每算出一组坐标就绘图的话会很慢，此处设置每算出多少次坐标后绘一次图。" & Chr(10) & Chr(10) & Chr(10) & "by Sykie Chen" & Chr(10) & "sykiechen@gmail.com" & Chr(10) & "项目详情请移至人人主页 陈仕玺.cmd 查看日志。", , "说明"
End Sub

Private Sub Command6_Click()
If (CLng(Text5.Text) / 10) >= 1 Then
    Text5.Text = CLng(Text5.Text) / 10
Else
    Text5.Text = "1"
End If
End Sub

Private Sub Command7_Click()
Text5.Text = CLng(Text5.Text) * 10
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
flag = True
End Sub

Private Sub Form_Resize()
Picture1.Width = Me.Width - Picture1.Left - 225
Picture1.Height = Me.Height - Picture1.Top - 525
End Sub
