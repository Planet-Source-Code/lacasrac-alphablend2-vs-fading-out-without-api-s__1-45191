VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alphablend_3 created by Laca--FADING OUT--"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fading out"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Felvesz1 
      Caption         =   "Create maps"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox SRC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   120
      Picture         =   "alpha3.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   0
      Width           =   3000
   End
   Begin VB.PictureBox SRC2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   120
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2865
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Alphablend project, created by Laca in 2003
'Kozari Laszlo, Hungary


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim rgbs()  As Long
Dim rgbs2() As Long
Dim cnt     As Byte
Dim ok      As Boolean
Private Sub Felvesz()

'This method is just create the 2
'picture maps by their rgb values.

ReDim rgbs(SRC.ScaleWidth - 1, SRC.ScaleHeight - 1)
ReDim rgbs2(SRC.ScaleWidth - 1, SRC.ScaleHeight - 1)

'A nagy
For X = 0 To SRC.ScaleWidth - 1
 For Y = 0 To SRC.ScaleHeight - 1
        C = GetPixel(SRC.hdc, X, Y)
        C2 = RGB(0, 0, 0) 'GetPixel(SRC2.hdc, X, Y)
                
        rgbs(X, Y) = C   'nagy
        rgbs2(X, Y) = C2 'kicsi
 Next Y
Next X

End Sub


Private Sub Command1_Click()

If Not ok Then Exit Sub
    Timer1.Interval = 1

End Sub


Private Sub Felvesz1_Click()

Call Felvesz
ok = True
Felvesz1.Enabled = Not ok

End Sub



Private Sub Made(Shade, x0, y0)
'Alphablend method
'Not too fast, but with the setpixel api _
 and with pixeldrawing its nice...

Alpha = Shade / 255
Alpha2 = (255 - Shade) / 255

bit0 = 255: bit1 = bit0 * 256: bit2 = bit1 * 256

SRC.Cls
For X = 0 To SRC2.ScaleWidth - 1
 For Y = 0 To SRC2.ScaleHeight - 1
      SRC1 = rgbs(X, Y): DST1 = rgbs2(X, Y)
      col = _
            (SRC1 And bit0) * Alpha + (DST1 And bit0) * Alpha2 Or _
            (SRC1 And bit1) * Alpha + (DST1 And bit1) * Alpha2 And bit1 Or _
            (SRC1 And bit2) * Alpha + (DST1 And bit2) * Alpha2 And bit2
      SetPixel SRC.hdc, X, Y, col
 Next Y
Next X
SRC.Refresh
    
End Sub






Private Sub Form_Activate()

SRC2.ScaleWidth = SRC.ScaleWidth
SRC2.ScaleHeight = SRC.ScaleHeight
cnt = 255

End Sub

Private Sub Timer1_Timer()

cnt = cnt - 15
If cnt <= 0 Then cnt = 0: Timer1.Interval = 0: Exit Sub

Call Made(cnt, 0, 0)
DoEvents

End Sub


