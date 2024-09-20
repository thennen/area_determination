VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Calibration"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   9345
   ClientWidth     =   3060
   LinkTopic       =   "Form3"
   ScaleHeight     =   3330
   ScaleWidth      =   3060
   Begin VB.CommandButton Command1 
      Caption         =   "Save Calibration"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox CalLength 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "10"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   $"Calibrate.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label NewCal 
      Caption         =   "New Calibration:  __ pixels/mm"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Calibration Length (mm)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Rpixel 
      Caption         =   "Right Pixel: __"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Lpixel 
      Caption         =   "Left Pixel: __"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CalLength_Change()
    If LeftPixelX <> 2000 And RightPixelX <> 2000 And LeftPixelX <> RightPixelX And CalLength.Text <> "" Then
        Form3.NewCal.Caption = "New Calibration: " & Round(Abs(RightPixelX - LeftPixelX) / CDbl(Form3.CalLength.Text), 2) & " pixels/mm"
    End If
End Sub

Private Sub Command1_Click()

        PixelsPerMM = Round(Abs(RightPixelX - LeftPixelX) / CDbl(Form3.CalLength.Text), 2)

            
        
            Open CALIBRATIONFILE For Output As #44
                  Print #44, PixelsPerMM; " Pixels/mm"
            Close #44
            
        Form1.CalLabel.Caption = "Current Calibration: " & PixelsPerMM & " pixels/mm"
        PixelArea = 1 / (PixelsPerMM * PixelsPerMM)
            
        

Form3.Hide
End Sub

Private Sub Form_Load()

End Sub
