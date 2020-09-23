VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alphablend_2 created by Laca"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Felvesz1 
      Caption         =   "Create maps"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   3000
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
      Left            =   1680
      Picture         =   "alpha2.frx":0000
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
      Height          =   1185
      Left            =   120
      Picture         =   "alpha2.frx":7485
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   0
      Top             =   0
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select Create maps,  then Move the mouse on the car's picture."
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
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
        C2 = GetPixel(SRC2.hdc, X, Y)
                
        rgbs(X, Y) = C   'nagy
        rgbs2(X, Y) = C2 'kicsi
 Next Y
Next X

'This line creates the alphablending...
Call Made(128, 0, 0) '50% fade

End Sub


Private Sub Felvesz1_Click()

Call Felvesz
ok = True
Felvesz1.Enabled = Not ok

End Sub

Private Sub SRC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not ok Then Exit Sub
    Call Made(128, X, Y)

End Sub


Private Sub v_Change()

Call Made(v.Value)
DST.Refresh

End Sub



Private Sub Made(Shade, x0, y0)
'Alphablend method
'Not too fast, but with the setpixel api _
 and with pixeldrawing its nice...
On Error Resume Next

Alpha = Shade / 255
Alpha2 = (255 - Shade) / 255

bit0 = 255
bit1 = bit0 * 256
bit2 = bit1 * 256
m_w = Int(SRC2.ScaleWidth / 2)
m_h = Int(SRC2.ScaleHeight / 2)


SRC.Cls
For X = 0 To SRC2.ScaleWidth - 1
 For Y = 0 To SRC2.ScaleHeight - 1
      SRC1 = rgbs((x0 - m_w) + X, (y0 - m_h) + Y): DST1 = rgbs2(X, Y)
      col = _
            (SRC1 And bit0) * Alpha + (DST1 And bit0) * Alpha2 Or _
            (SRC1 And bit1) * Alpha + (DST1 And bit1) * Alpha2 And bit1 Or _
            (SRC1 And bit2) * Alpha + (DST1 And bit2) * Alpha2 And bit2
      SetPixel SRC.hdc, (x0 - m_w) + X, (y0 - m_h) + Y, col
 Next Y
Next X
SRC.Refresh
    
End Sub

Private Sub v_Scroll()
v_Change
End Sub





