VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Video"
   ClientHeight    =   11535
   ClientLeft      =   3195
   ClientTop       =   300
   ClientWidth     =   18255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11535
   ScaleWidth      =   18255
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar HScroll1 
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   11280
      Width           =   17895
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   11535
      Left            =   18000
      TabIndex        =   2
      Top             =   0
      Width           =   195
   End
   Begin VB.PictureBox picWindow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   12480
      Left            =   -120
      ScaleHeight     =   12420
      ScaleWidth      =   18075
      TabIndex        =   0
      Top             =   -960
      Width           =   18135
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   15540
         Left            =   120
         ScaleHeight     =   15480
         ScaleWidth      =   19515
         TabIndex        =   1
         Top             =   960
         Width           =   19575
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub Form_Activate()

    Picture1.SetFocus
    
End Sub
Private Sub Form_Load()

    picWindow.Left = 0
    picWindow.Top = 0
    'picWindow.Width = Me.Width - VScroll1.Width
    'picWindow.Height = Me.Height - 300
    
    VScroll1.SmallChange = 100
    VScroll1.LargeChange = picWindow.Height
    VScroll1.Top = 0
    
    HScroll1.SmallChange = 100
    HScroll1.LargeChange = picWindow.Width
    HScroll1.Left = 0
    
    Picture1.Left = 0
    Picture1.Top = 0

End Sub

Public Sub CheckKeyCode(KeyCode As Integer)

    Dim nScrollValue As Double
    Dim nOnePage As Integer
    
    nOnePage = Me.VScroll1.Height
    
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If KeyCode = vbKeyPageDown Then
            nScrollValue = -Me.Picture1.Top + nOnePage
        Else
            nScrollValue = -Me.Picture1.Top - nOnePage
        End If
        'Make sure that the scroll bar 'handle' is not attempted to be positioned
        'below the bottom of the scroll bar.
        If nScrollValue > Me.VScroll1.Max Then
            nScrollValue = Me.VScroll1.Max
            Me.Picture.Top = -Me.VScroll1.Max
        End If
        If nScrollValue > 0 Then
            Me.VScroll1.Value = nScrollValue
        Else
            Me.VScroll1.Value = 0
        End If
    End If
    
End Sub


Private Sub Form_Resize()
    picWindow.Width = Me.Width - 100
    picWindow.Height = Me.Height - 350
    VScroll1.Left = Me.Width - VScroll1.Width - 100
    
    VScroll1.Height = picWindow.Height - 200
    VScroll1.Max = Picture1.Height - picWindow.Height + 200
    
    HScroll1.Top = Me.Height - VScroll1.Width - 340
    HScroll1.Width = picWindow.Width - 200
    HScroll1.Max = Picture1.Width - picWindow.Width + 200
    
    If Picture1.Height <= Me.Height - 550 Then
        VScroll1.Visible = False
    Else
        VScroll1.Visible = True
    End If
    
    If Picture1.Width <= Me.Width - 300 Then
        HScroll1.Visible = False
    Else
        HScroll1.Visible = True
    End If

    
End Sub

Private Sub picture1_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckKeyCode KeyCode
    
End Sub

Private Sub VScroll1_Change()

    Picture1.Top = -(VScroll1.Value)

    ' So that the scrollbar doesn't blink
    Picture1.SetFocus

End Sub

Private Sub VScroll1_Scroll()
    Picture1.Top = -(VScroll1.Value)
    Picture1.SetFocus
End Sub
Private Sub HScroll1_Change()

    Picture1.Left = -(HScroll1.Value)

    ' So that the scrollbar doesn't blink
    Picture1.SetFocus

End Sub

Private Sub HScroll1_Scroll()
    Picture1.Left = -(HScroll1.Value)
    Picture1.SetFocus
End Sub


Private Sub Picture1_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Form3.Visible = True Then

    If Button = 1 Then
        LeftPixelX = X / 15
        Form3.Lpixel.Caption = "Left Pixel: " & LeftPixelX
    End If
    If Button = 2 Then
        RightPixelX = X / 15
        Form3.Rpixel.Caption = "Right Pixel: " & RightPixelX
    End If

    If LeftPixelX <> 2000 And RightPixelX <> 2000 And LeftPixelX <> RightPixelX Then
        PixelsPerMM = Round(Abs(RightPixelX - LeftPixelX) / CDbl(Form3.CalLength.Text), 2)
        Form3.NewCal.Caption = "New Calibration: " & PixelsPerMM & " pixels/mm"
    End If
    
End If



End Sub
