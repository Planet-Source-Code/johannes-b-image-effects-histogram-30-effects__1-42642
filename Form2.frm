VERSION 5.00
Begin VB.Form Histogram 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Histogram"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4065
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   3840
      Width           =   2175
      Begin VB.Label Label3 
         Caption         =   "Level: 0"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Quantity: 0"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   2040
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   5
      Left            =   1800
      Max             =   20
      Min             =   1
      TabIndex        =   7
      Top             =   3480
      Value           =   1
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Make histogram with all channels"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Make histogram with blue channel"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Make histogram with green channel"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make histogram with red channel"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3735
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   2
      Top             =   1800
      Width           =   3840
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   120
      Width           =   3840
   End
   Begin VB.Label Label1 
      Caption         =   "Skip pixles (faster)"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "Histogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Quality As Byte

Dim Rred
Dim Ggreen
Dim Bblue

Dim Max

Private Sub CalculateMax()
Max = 0
For Counter = 0 To 254
If Val(List1.List(Counter)) > Max Then Max = List1.List(Counter)
Next
Max = Max / Picture1.ScaleHeight
End Sub

Private Sub DrawHistogram()
Picture1.ForeColor = vbWhite
For Counter = 0 To 254
Picture1.Line (Counter, Picture1.ScaleHeight - (List1.List(Counter) / Max))-(Counter, 0)
Next
Picture1.Refresh
End Sub


Private Sub GetRGB(ByVal Col As String)
On Error Resume Next
    Bblue = Col \ (256 ^ 2)
    Ggreen = (Col - Bblue * 256 ^ 2) \ 256
    Rred = (Col - Bblue * 256 ^ 2 - Ggreen * 256) '\ 256
End Sub
Private Sub ShowGradientRed()
For Counter = 0 To 254
Picture2.ForeColor = RGB(Counter, 0, 0)
Picture2.Line (Counter, Picture2.ScaleHeight)-(Counter, 0)
Next
Picture2.Refresh
End Sub

Private Sub ShowGradientGreen()
For Counter = 0 To 254
Picture2.ForeColor = RGB(0, Counter, 0)
Picture2.Line (Counter, Picture2.ScaleHeight)-(Counter, 0)
Next
Picture2.Refresh
End Sub
Private Sub ShowGradientBlue()
For Counter = 0 To 254
Picture2.ForeColor = RGB(0, 0, Counter)
Picture2.Line (Counter, Picture2.ScaleHeight)-(Counter, 0)
Next
Picture2.Refresh
End Sub
Private Sub ShowGradientAll()
For Counter = 0 To 254
Picture2.ForeColor = RGB(Counter, Counter, Counter)
Picture2.Line (Counter, Picture2.ScaleHeight)-(Counter, 0)
Next
Picture2.Refresh
End Sub
Private Sub Command1_Click()
Screen.MousePointer = "11"
Quality = HScroll1.Value
List1.Clear
Picture1.Cls
For Counter = 0 To 254
List1.AddItem "0"
Next
For YYY = 0 To Form1.Picture1.ScaleHeight Step Quality
For XXX = 0 To Form1.Picture1.ScaleWidth Step Quality
Pixel = GetPixel(Form1.Picture1.HDC, XXX, YYY)
GetRGB Pixel
List1.List(Rred) = Val(List1.List(Rred)) + 1
Next
Next
ShowGradientRed
CalculateMax
DrawHistogram
Screen.MousePointer = "0"
End Sub

Private Sub Command2_Click()
Screen.MousePointer = "11"
Quality = HScroll1.Value
List1.Clear
Picture1.Cls
For Counter = 0 To 254
List1.AddItem "0"
Next
For YYY = 0 To Form1.Picture1.ScaleHeight Step Quality
For XXX = 0 To Form1.Picture1.ScaleWidth Step Quality
Pixel = GetPixel(Form1.Picture1.HDC, XXX, YYY)
GetRGB Pixel
List1.List(Ggreen) = Val(List1.List(Ggreen)) + 1
Next
Next
ShowGradientGreen
CalculateMax
DrawHistogram
Screen.MousePointer = "0"
End Sub


Private Sub Command3_Click()
Screen.MousePointer = "11"
Quality = HScroll1.Value
List1.Clear
Picture1.Cls
For Counter = 0 To 254
List1.AddItem "0"
Next
For YYY = 0 To Form1.Picture1.ScaleHeight Step Quality
For XXX = 0 To Form1.Picture1.ScaleWidth Step Quality
Pixel = GetPixel(Form1.Picture1.HDC, XXX, YYY)
GetRGB Pixel
List1.List(Bblue) = Val(List1.List(Bblue)) + 1
Next
Next
ShowGradientBlue
CalculateMax
DrawHistogram
Screen.MousePointer = "0"
End Sub


Private Sub Command4_Click()
Screen.MousePointer = "11"
Quality = HScroll1.Value
List1.Clear
Picture1.Cls
For Counter = 0 To 254
List1.AddItem "0"
Next
For YYY = 0 To Form1.Picture1.ScaleHeight Step Quality
For XXX = 0 To Form1.Picture1.ScaleWidth Step Quality
Pixel = GetPixel(Form1.Picture1.HDC, XXX, YYY)
GetRGB Pixel
List1.List((Rred + Ggreen + Bblue) / 3) = Val(List1.List((Rred + Ggreen + Bblue) / 3)) + 1
Next
Next
ShowGradientAll
CalculateMax
DrawHistogram
Screen.MousePointer = "0"
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "Quantity: " & List1.List(X)
Label3.Caption = "Level: " & X
End Sub


