VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Effects version 1.0 by Johannes B 2002"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   614
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command35 
      Caption         =   "C"
      Height          =   255
      Left            =   3720
      TabIndex        =   45
      ToolTipText     =   "Center scroll"
      Top             =   5400
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4335
      Left            =   3720
      TabIndex        =   44
      Top             =   1080
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   5400
      Width           =   3615
   End
   Begin VB.PictureBox PC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   41
      Top             =   1080
      Width           =   3615
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3960
         Left            =   0
         Picture         =   "Form1.frx":08CA
         ScaleHeight     =   264
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   221
         TabIndex        =   42
         Top             =   0
         Width           =   3315
      End
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Modern art"
      Height          =   255
      Left            =   4200
      TabIndex        =   40
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Strange..."
      Height          =   255
      Left            =   4200
      TabIndex        =   36
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Y blur"
      Height          =   255
      Left            =   6720
      TabIndex        =   35
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X blur"
      Height          =   255
      Left            =   6720
      TabIndex        =   34
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Darkness..."
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Histogram..."
      Height          =   255
      Left            =   4200
      TabIndex        =   33
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Lightness..."
      Height          =   255
      Left            =   4200
      TabIndex        =   32
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton Command29 
      Caption         =   "3D effect..."
      Height          =   255
      Left            =   4200
      TabIndex        =   30
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Colorize..."
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command27 
      Caption         =   "3D grid..."
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Grayscale..."
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Save picture..."
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Open picture..."
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   360
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   2880
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Fog..."
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Find vertical edges b/w..."
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Find horizontal edges b/w..."
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Find vertical edges..."
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Find horizontal edges..."
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Smart noise..."
      Height          =   255
      Left            =   6720
      TabIndex        =   21
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Black and white"
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Pixelize..."
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Add noise..."
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Flip vertical"
      Height          =   255
      Left            =   6720
      TabIndex        =   17
      Top             =   5160
      Width           =   2175
   End
   Begin VB.PictureBox TempPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3600
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   16
      Top             =   600
      Width           =   375
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Flip horizontal"
      Height          =   255
      Left            =   6720
      TabIndex        =   15
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Mirror down in top"
      Height          =   255
      Left            =   6720
      TabIndex        =   14
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Mirror top in down"
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Mirror right in left"
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Mirror left in right"
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Invert"
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Darken..."
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Lighten..."
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.CheckBox Cblue 
      Caption         =   "Blue"
      Height          =   255
      Left            =   8280
      TabIndex        =   5
      Top             =   3240
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox Cgreen 
      Caption         =   "Green"
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   3240
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox Cred 
      Caption         =   "Red"
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   3240
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Restore image"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   3855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Divide colors..."
      Height          =   255
      Left            =   6720
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Multiply colors..."
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape Shape5 
      Height          =   735
      Left            =   120
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
      Height          =   255
      Left            =   4080
      TabIndex        =   39
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Effects (with channel selector)"
      Height          =   255
      Left            =   6600
      TabIndex        =   38
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Effects"
      Height          =   255
      Left            =   4080
      TabIndex        =   37
      Top             =   0
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      Height          =   495
      Left            =   4080
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      Height          =   4695
      Left            =   4080
      Top             =   240
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      Height          =   1815
      Left            =   6600
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   6600
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pixel
Dim Pixel2

Dim Rred
Dim Ggreen
Dim Bblue

Dim RR1
Dim GG1
Dim BB1

Dim RR2
Dim GG2
Dim BB2

Dim RR3
Dim GG3
Dim BB3

Dim Q As String
Dim Q2 As String

Dim Temp As Integer
Dim Temp2 As Integer

Dim XXX As Integer
Dim YYY As Integer

Dim XX As Integer
Dim YY As Integer

Dim RR As Integer
Dim RG As Integer
Dim RB As Integer

Dim CurX
Dim CurY

Dim JB As Byte

Sub UppdateScroll()
If Picture1.Width <= PC.ScaleWidth Then
HScroll1.Enabled = False
Else
HScroll1.Enabled = True
End If

If Picture1.Height <= PC.ScaleHeight Then
VScroll1.Enabled = False
Else
VScroll1.Enabled = True
End If

VScroll1.Max = Picture1.ScaleHeight - PC.ScaleHeight
HScroll1.Max = Picture1.ScaleWidth - PC.ScaleWidth

If HScroll1.Enabled = False Then
Picture1.Left = (PC.Width / 2) - (Picture1.Width / 2)
End If
If VScroll1.Enabled = False Then
Picture1.Top = (PC.Height / 2) - (Picture1.Height / 2)
End If

HScroll1.Value = 0
VScroll1.Value = 0
End Sub




Private Sub GetRGB(ByVal Col As String)
On Error Resume Next
    Bblue = Col \ (256 ^ 2)
    Ggreen = (Col - Bblue * 256 ^ 2) \ 256
    Rred = (Col - Bblue * 256 ^ 2 - Ggreen * 256) '\ 256
End Sub
Private Sub Command1_Click()
On Error Resume Next
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel
RR1 = Rred
GG1 = Ggreen
BB1 = Bblue
If XXX < Picture1.ScaleWidth - 3 Then Pixel2 = GetPixel(Picture1.HDC, XXX + 2, YYY)
GetRGB Pixel2
RR2 = Rred
GG2 = Ggreen
BB2 = Bblue
If Cred.Value = 1 Then Rred = (RR1 + RR2) / 2
If Cgreen.Value = 1 Then Ggreen = (GG1 + GG2) / 2
If Cblue.Value = 1 Then Bblue = (BB1 + BB2) / 2
SetPixelV Picture1.HDC, XXX + 1, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command10_Click()
On Error Resume Next
Q = InputBox("Enter a value for find vertical edges b/w (0-255, higher value = less edges)", "", "7")
If Q = "" Then Exit Sub

For XXX = 0 To Picture1.ScaleWidth - 1
For YYY = 0 To Picture1.ScaleHeight - 1

Pixel2 = GetPixel(Picture1.HDC, XXX, YYY + 2)
Pixel = GetPixel(Picture1.HDC, XXX, YYY + 1)

GetRGB Pixel
RR1 = Rred
GG1 = Ggreen
BB1 = Bblue

GetRGB Pixel2
RR2 = Rred
GG2 = Ggreen
BB2 = Bblue

Temp = (RR1 + GG1 + BB1)
Temp = (Temp / 3)

Temp2 = (RR2 + GG2 + BB2)
Temp2 = (Temp2 / 3)

If Temp = Temp2 Then Pixel = vbWhite
If Val(Temp) > Val(Temp2) Then
If Val(Temp) - Val(Temp2) >= Q Then
Pixel = vbBlack
Else
Pixel = vbWhite
End If
Else
If Val(Temp2) - Val(Temp) >= Q Then
Pixel = vbBlack
Else
Pixel = vbWhite
End If
End If


SetPixelV Picture1.HDC, XXX, YYY, Pixel
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command11_Click()
On Error Resume Next
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To (Picture1.ScaleWidth / 2) - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
SetPixelV Picture1.HDC, Picture1.ScaleWidth - XXX, YYY, Pixel
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command12_Click()
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To (Picture1.ScaleWidth / 2) - 1
Pixel = GetPixel(Picture1.HDC, Picture1.ScaleWidth - XXX, YYY)
SetPixelV Picture1.HDC, XXX, YYY, Pixel
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub


Private Sub Command13_Click()
On Error Resume Next
For YYY = 0 To (Picture1.ScaleHeight / 2) - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
SetPixelV Picture1.HDC, XXX, Picture1.ScaleHeight - YYY, Pixel
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command14_Click()
On Error Resume Next
For YYY = 0 To (Picture1.ScaleHeight / 2) - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, Picture1.ScaleHeight - YYY)
SetPixelV Picture1.HDC, XXX, YYY, Pixel
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub


Private Sub Command15_Click()
On Error Resume Next
TempPic.Width = Picture1.Width
TempPic.Height = Picture1.Height
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
SetPixelV TempPic.HDC, Picture1.ScaleWidth - (XXX + 1), YYY, Pixel
Next
Picture1.Refresh
Next
BitBlt Picture1.HDC, 0, 0, TempPic.ScaleWidth - 1, TempPic.ScaleHeight - 1, TempPic.HDC, 0, 0, &HCC0020

Picture1.Refresh
End Sub

Private Sub Command16_Click()
On Error Resume Next
TempPic.Width = Picture1.Width
TempPic.Height = Picture1.Height
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
SetPixelV TempPic.HDC, XXX, Picture1.ScaleHeight - (YYY + 1), Pixel
Next
Picture1.Refresh
Next
BitBlt Picture1.HDC, 0, 0, TempPic.ScaleWidth - 1, TempPic.ScaleHeight - 1, TempPic.HDC, 0, 0, &HCC0020

Picture1.Refresh
End Sub


Private Sub Command17_Click()
Randomize
On Error Resume Next
Q2 = InputBox("Select mode (0 = use image colors, 1 = random colors)", "", "0")
If Q2 = "" Then Exit Sub
Q = InputBox("Enter a value for add noise", "", "20000")
If Q = "" Then Exit Sub

For XXX = 0 To Q
If Q2 = "0" Then
Pixel = GetPixel(Picture1.HDC, Rnd * Picture1.ScaleWidth, Rnd * Picture1.ScaleHeight)
Else
Pixel = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End If
SetPixelV Picture1.HDC, Rnd * Picture1.ScaleWidth, Rnd * Picture1.ScaleHeight, Pixel
Next
Picture1.Refresh
End Sub

Private Sub Command18_Click()
On Error Resume Next
On Error Resume Next
Q = InputBox("Enter a value for pixelize", "", "5")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1 Step Q
For XXX = 0 To Picture1.ScaleWidth - 1 Step Q
Pixel = GetPixel(Picture1.HDC, XXX + 1, YYY + 1)
Picture1.Line (XXX, YYY)-(XXX + Q, YYY + Q), Pixel, BF

Next
Picture1.Refresh
Next
Picture1.Refresh

End Sub


Private Sub Command19_Click()
On Error Resume Next
Q = InputBox("Enter a value for black and white (0-255, high value will make a darker image)", "", "127")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1

Pixel = GetPixel(Picture1.HDC, XXX, YYY)

GetRGB Pixel

Temp = (Rred + Ggreen + Bblue)
Temp = (Temp / 3)

If Val(Temp) >= Q Then
Pixel = vbWhite
Else
Pixel = vbBlack
End If




SetPixelV Picture1.HDC, XXX, YYY, Pixel
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command2_Click()
On Error Resume Next

For XXX = 0 To Picture1.ScaleWidth - 1
For YYY = 0 To Picture1.ScaleHeight - 1

Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel
RR1 = Rred
GG1 = Ggreen
BB1 = Bblue
If YYY < Picture1.ScaleHeight - 3 Then Pixel2 = GetPixel(Picture1.HDC, XXX, YYY + 2)
GetRGB Pixel2
RR2 = Rred
GG2 = Ggreen
BB2 = Bblue
If Cred.Value = 1 Then Rred = (RR1 + RR2) / 2
If Cgreen.Value = 1 Then Ggreen = (GG1 + GG2) / 2
If Cblue.Value = 1 Then Bblue = (BB1 + BB2) / 2
SetPixelV Picture1.HDC, XXX, YYY + 1, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub


Private Sub Command20_Click()
Randomize
On Error Resume Next
Q = InputBox("Enter a value for smart noise (0-255)", "", "50")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

RR = Rnd * 1
RG = Rnd * 1
RB = Rnd * 1

If RR = 1 Then
If Cred.Value = 1 Then Rred = Rred + Rnd * Q
Else
If Cred.Value = 1 Then Rred = Rred - Rnd * Q
End If
If RG = 1 Then
If Cgreen.Value = 1 Then Ggreen = Ggreen + Rnd * Q
Else
If Cgreen.Value = 1 Then Ggreen = Ggreen - Rnd * Q
End If
If RB = 1 Then
If Cblue.Value = 1 Then Bblue = Bblue + Rnd * Q
Else
If Cblue.Value = 1 Then Bblue = Bblue - Rnd * Q
End If

SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command21_Click()
On Error Resume Next
Q = InputBox("Enter a value for find horizontal edges (higher value = brighter image)", "", "4")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1

Pixel2 = GetPixel(Picture1.HDC, XXX + 2, YYY)
Pixel = GetPixel(Picture1.HDC, XXX + 1, YYY)

GetRGB Pixel
RR1 = Rred
GG1 = Ggreen
BB1 = Bblue

GetRGB Pixel2
RR2 = Rred
GG2 = Ggreen
BB2 = Bblue

If RR1 = RR2 Then RR3 = 0
If RR1 > RR2 Then
RR3 = RR1 - RR2
Else
RR3 = RR2 - RR1
End If

If GG1 = GG2 Then GG3 = 0
If GG1 > GG2 Then
GG3 = GG1 - GG2
Else
GG3 = GG2 - GG1
End If

If BB1 = BB2 Then BB3 = 0
If BB1 > BB2 Then
BB3 = BB1 - BB2
Else
BB3 = BB2 - BB1
End If

SetPixelV Picture1.HDC, XXX, YYY, RGB(RR3 * Q, GG3 * Q, BB3 * Q)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command22_Click()
On Error Resume Next
Q = InputBox("Enter a value for find vertical edges (higher value = brighter image)", "", "4")
If Q = "" Then Exit Sub

For XXX = 0 To Picture1.ScaleWidth - 1
For YYY = 0 To Picture1.ScaleHeight - 1

Pixel2 = GetPixel(Picture1.HDC, XXX, YYY + 2)
Pixel = GetPixel(Picture1.HDC, XXX, YYY + 1)

GetRGB Pixel
RR1 = Rred
GG1 = Ggreen
BB1 = Bblue

GetRGB Pixel2
RR2 = Rred
GG2 = Ggreen
BB2 = Bblue

If RR1 = RR2 Then RR3 = 0
If RR1 > RR2 Then
RR3 = RR1 - RR2
Else
RR3 = RR2 - RR1
End If

If GG1 = GG2 Then GG3 = 0
If GG1 > GG2 Then
GG3 = GG1 - GG2
Else
GG3 = GG2 - GG1
End If

If BB1 = BB2 Then BB3 = 0
If BB1 > BB2 Then
BB3 = BB1 - BB2
Else
BB3 = BB2 - BB1
End If


SetPixelV Picture1.HDC, XXX, YYY, RGB(RR3 * Q, GG3 * Q, BB3 * Q)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command23_Click()
On Error Resume Next
Q = InputBox("Enter a value for fog", "", "30")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Cred.Value = 1 Then
If Val(Rred) > 127 Then
Rred = Rred - Q
If Rred < 127 Then Rred = 127
Else
Rred = Rred + Q
If Rred > 127 Then Rred = 127
End If
End If
If Cgreen.Value = 1 Then
If Val(Ggreen) > 127 Then
Ggreen = Ggreen - Q
If Ggreen < 127 Then Ggreen = 127
Else
Ggreen = Ggreen + Q
If Ggreen > 127 Then Ggreen = 127
End If
End If


If Cblue.Value = 1 Then
If Val(Bblue) > 127 Then
Bblue = Bblue - Q
If Bblue < 127 Then Bblue = 127
Else
Bblue = Bblue + Q
If Bblue > 127 Then Bblue = 127
End If
End If

SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command24_Click()
CM.CancelError = True
On Error GoTo ja
CM.Filter = "Image|*.bmp;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur"
CM.ShowOpen
Picture1.Picture = LoadPicture(CM.FileName)
UppdateScroll
Exit Sub
ja:
Exit Sub
End Sub


Private Sub Command25_Click()
CM.CancelError = True
On Error GoTo ja
CM.Filter = "Bitmap|*.bmp"
CM.ShowSave
SavePicture Picture1.Image, CM.FileName
Exit Sub
ja:
Exit Sub
End Sub





Private Sub Command26_Click()
On Error Resume Next
Q = InputBox("Channels to read from? (0 = All, 1 = Red, 2 = Green, 3 = Blue)", "", "0")
If Q = "" Then Exit Sub
If Q > 3 Then Exit Sub
If Q < 0 Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1

Pixel = GetPixel(Picture1.HDC, XXX, YYY)

GetRGB Pixel

If Q = 0 Then
Temp = (Rred + Ggreen + Bblue)
Temp = (Temp / 3)
End If
If Q = 1 Then
Temp = (Rred)
End If
If Q = 2 Then
Temp = (Ggreen)
End If
If Q = 3 Then
Temp = (Bblue)
End If


SetPixelV Picture1.HDC, XXX, YYY, RGB(Temp, Temp, Temp)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub





Private Sub Command27_Click()
On Error Resume Next
Q2 = InputBox("Enter number of steps for 3D grid", "", "4")
If Q2 = "" Then Exit Sub
Q = InputBox("Enter a brightness value for 3D grid (higher value = darker image)", "", "10")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1 Step Q2 + 1
For XXX = 0 To Picture1.ScaleWidth - 1 Step Q2 + 1

Pixel = GetPixel(Picture1.HDC, XXX, YYY)

GetRGB Pixel
Rred = Rred - Q
Ggreen = Ggreen - Q
Bblue = Bblue - Q


For Counter = 1 To Q2
SetPixelV Picture1.HDC, XXX + Counter, YYY, RGB(Rred, Ggreen, Bblue)
Next
For Counter = 1 To Q2
SetPixelV Picture1.HDC, XXX - Counter, YYY, RGB(Rred, Ggreen, Bblue)
Next
For Counter = 1 To Q2
SetPixelV Picture1.HDC, XXX, YYY + Counter, RGB(Rred, Ggreen, Bblue)
Next
For Counter = 1 To Q2
SetPixelV Picture1.HDC, XXX, YYY - Counter, RGB(Rred, Ggreen, Bblue)
Next

Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub





Private Sub Command28_Click()
On Error GoTo ja

CM.CancelError = True
CM.ShowColor
GetRGB CM.Color
RR3 = Rred
GG3 = Ggreen
BB3 = Bblue
On Error Resume Next
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

Temp = (Rred + Ggreen + Bblue)
Temp = Temp / 3

SetPixelV Picture1.HDC, XXX, YYY, RGB((RR3 + Temp), (GG3 + Temp), (BB3 + Temp))
Next
Picture1.Refresh
Next
Picture1.Refresh
Exit Sub
ja:
Exit Sub
End Sub

Private Sub Command29_Click()
On Error Resume Next
Q = InputBox("Enter depth of 3D convertion (lower number = deeper)", "", "6")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

Temp = (Rred + Ggreen + Bblue)
Temp = Temp / 3

Picture1.ForeColor = RGB(Rred, Ggreen, Bblue)
Picture1.Line (XXX, YYY)-(XXX, YYY - (Temp / Q))
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command3_Click()
On Error Resume Next
Q = InputBox("Enter a value for divine colors", "", "1,5")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Cred.Value = 1 Then Rred = Rred / Q
If Cgreen.Value = 1 Then Ggreen = Ggreen / Q
If Cblue.Value = 1 Then Bblue = Bblue / Q

SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command30_Click()
On Error Resume Next
Q = InputBox("Enter a value for darkness", "", "1,5")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Rred > 128 Then
RR1 = Rred - 128
Else
RR1 = 128 - Rred
End If

If Ggreen > 128 Then
GG1 = Ggreen - 128
Else
GG1 = 128 - Ggreen
End If

If Bblue > 128 Then
BB1 = Bblue - 128
Else
BB1 = 128 - Bblue
End If

RR1 = RR1 / Q
GG1 = GG1 / Q
BB1 = BB1 / Q

SetPixelV Picture1.HDC, XXX, YYY, RGB(RR1, GG1, BB1)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command31_Click()
On Error Resume Next
Q = InputBox("Enter a value for lightness", "", "3")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Rred > 128 Then
RR1 = Rred - 128
Else
RR1 = 128 - Rred
End If

If Ggreen > 128 Then
GG1 = Ggreen - 128
Else
GG1 = 128 - Ggreen
End If

If Bblue > 128 Then
BB1 = Bblue - 128
Else
BB1 = 128 - Bblue
End If

RR1 = RR1 * Q
GG1 = GG1 * Q
BB1 = BB1 * Q

SetPixelV Picture1.HDC, XXX, YYY, RGB(RR1, GG1, BB1)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub


Private Sub Command32_Click()
Histogram.Show
End Sub





Private Sub Command33_Click()
On Error Resume Next
Q = InputBox("Enter a value for strange", "", "1,1")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Rred > 127 Then
Rred = 255 - Rred / Q
Else
Rred = 0 + Rred / Q
End If

If Ggreen > 127 Then
Ggreen = 255 - Ggreen / Q
Else
Ggreen = 0 + Ggreen / Q
End If

If Bblue > 127 Then
Bblue = 255 - Bblue / Q
Else
Bblue = 0 + Bblue / Q
End If

SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command34_Click()
On Error Resume Next

For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Rred > 127 Then
RR1 = 255 - Rred
Rred = Rred / RR1
Else
RR1 = 255 - Rred
Rred = Rred * RR1
End If

If Ggreen > 127 Then
GG1 = 255 - Ggreen
Ggreen = Ggreen / GG1
Else
GG1 = 255 - Ggreen
Ggreen = Ggreen * GG1
End If

If Bblue > 127 Then
BB1 = 255 - Bblue
Bblue = Bblue / BB1
Else
BB1 = 255 - Bblue
Bblue = Bblue * BB1
End If

SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub


Private Sub Command35_Click()
If HScroll1.Enabled = True Then HScroll1.Value = HScroll1.Max / 2
If VScroll1.Enabled = True Then VScroll1.Value = VScroll1.Max / 2
End Sub

Private Sub Command36_Click()

End Sub

Private Sub Command4_Click()
On Error Resume Next
Q = InputBox("Enter a value for multiply colors", "", "1,5")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Cred.Value = 1 Then Rred = Rred * Q
If Cgreen.Value = 1 Then Ggreen = Ggreen * Q
If Cblue.Value = 1 Then Bblue = Bblue * Q

SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub


Private Sub Command5_Click()
Picture1.Cls
End Sub


Private Sub Command6_Click()
On Error Resume Next
Q = InputBox("Enter a value for darken", "", "30")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Cred.Value = 1 Then Rred = Rred - Q
If Cgreen.Value = 1 Then Ggreen = Ggreen - Q
If Cblue.Value = 1 Then Bblue = Bblue - Q

SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub

Private Sub Command7_Click()
On Error Resume Next
Q = InputBox("Enter a value for lighten", "", "30")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Cred.Value = 1 Then Rred = Rred + Q
If Cgreen.Value = 1 Then Ggreen = Ggreen + Q
If Cblue.Value = 1 Then Bblue = Bblue + Q

SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub


Private Sub Command8_Click()
On Error Resume Next
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1
Pixel = GetPixel(Picture1.HDC, XXX, YYY)
GetRGB Pixel

If Cred.Value = 1 Then Rred = 255 - Rred
If Cgreen.Value = 1 Then Ggreen = 255 - Ggreen
If Cblue.Value = 1 Then Bblue = 255 - Bblue

SetPixelV Picture1.HDC, XXX, YYY, RGB(Rred, Ggreen, Bblue)
Next
Picture1.Refresh
Next
Picture1.Refresh


End Sub


Private Sub Command9_Click()
On Error Resume Next
Q = InputBox("Enter a value for find horizontal edges b/w (0-255, higher value = less edges)", "", "7")
If Q = "" Then Exit Sub
For YYY = 0 To Picture1.ScaleHeight - 1
For XXX = 0 To Picture1.ScaleWidth - 1

Pixel2 = GetPixel(Picture1.HDC, XXX + 2, YYY)
Pixel = GetPixel(Picture1.HDC, XXX + 1, YYY)

GetRGB Pixel
RR1 = Rred
GG1 = Ggreen
BB1 = Bblue

GetRGB Pixel2
RR2 = Rred
GG2 = Ggreen
BB2 = Bblue

Temp = (RR1 + GG1 + BB1)
Temp = (Temp / 3)

Temp2 = (RR2 + GG2 + BB2)
Temp2 = (Temp2 / 3)

If Temp = Temp2 Then Pixel = vbWhite
If Val(Temp) > Val(Temp2) Then
If Val(Temp) - Val(Temp2) >= Q Then
Pixel = vbBlack
Else
Pixel = vbWhite
End If
Else
If Val(Temp2) - Val(Temp) >= Q Then
Pixel = vbBlack
Else
Pixel = vbWhite
End If
End If


SetPixelV Picture1.HDC, XXX, YYY, Pixel
Next
Picture1.Refresh
Next
Picture1.Refresh
End Sub


Private Sub Form_Load()
UppdateScroll
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Thanks for downloading my code! Please vote or leave a comment if you liked it!", vbInformation
End Sub

Private Sub HScroll1_Change()
Picture1.Left = 0 - HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
JB = 1
CurX = X
CurY = Y
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If JB = 1 Then
If HScroll1.Enabled = True Then HScroll1.Value = HScroll1.Value + CurX - X
If VScroll1.Enabled = True Then VScroll1.Value = VScroll1.Value + CurY - Y
End If
End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
JB = 0
End Sub


Private Sub VScroll1_Change()
Picture1.Top = 0 - VScroll1.Value
End Sub


Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub


