Attribute VB_Name = "Module1"
Option Explicit
Public FrameCount As Long
Dim ImageBuf() As Byte
Public b() As Byte, blacklevel As Byte
Dim Row As Long, Column As Long
Public LeftPixelX As Long
Public RightPixelX As Long
Public PixelsPerMM As Double
Public PixelArea As Double

Public Const CALIBRATIONFILE As String = "calibration.txt"

Public Const RAWDATA_IMAGE = 0
Public Const BMPDATA_IMAGE = 1
Public Const NORMAL_MODE = 0
Public Const TRIGGER_MODE = 1
Public Const GRAB_FRAME_FOREVER = &H8888


Public Type TPROCESSEDDATAPROPERTY
    CameraID As Long
    Row As Long
    Column As Long
    Bin As Long
    XStart As Long
    YStart As Long
    exposureTime As Long
    RedGain As Long
    GreenGain As Long
    BlueGain As Long
    TimeStamp As Long
    TriggerOccurred As Long
    TriggerEventCount As Long
    ProcessFrameType As Long
End Type


Declare Function BUFUSB_InitDevice Lib "BUF_USBCamera_SDK_Stdcall" () As Long
Declare Function BUFUSB_UnInitDevice Lib "BUF_USBCamera_SDK_Stdcall" () As Long
Declare Function BUFUSB_GetModuleNoSerialNo Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal ModuleNo As String, ByVal SerialNo As String) As Long
Declare Function BUFUSB_AddDeviceToWorkingSet Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long) As Long
Declare Function BUFUSB_RemoveDeviceFromWorkingSet Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long) As Long
Declare Function BUFUSB_StartCameraEngine Lib "BUF_USBCamera_SDK_Stdcall" (ByVal ParentHandle As Long, ByVal CameraBitOption As Long) As Long
Declare Function BUFUSB_StopCameraEngine Lib "BUF_USBCamera_SDK_Stdcall" () As Long
Declare Function BUFUSB_SetCameraWorkMode Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal WorkMode As Long) As Long
' The following two APIs are not recommended...we recommend user to install frame callback for NORMAL and TRIGGER modes
'Declare Function MTUSB_SetExternalParameters Lib "MT_USBCamera_SDK_Stdcall" ( ByVal DevHandle As Long, ByVal AutoLoop As Byte, ByVal IsRawGraph as Byte, ByVal IsJPEG As Byte, ByVal FilePath as String, ByVal FileName As String) as long
'Declare Function MTUSB_WaitingExternalTrigger Lib "MT_USBCamera_SDK_Stdcall" ( DEV_HANDLE DevHandle, bool StartWait, CallBackFunc Aproc );
Declare Function BUFUSB_ShowFactoryControlPanel Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal passWord As String) As Long
Declare Function BUFUSB_HideFactoryControlPanel Lib "BUF_USBCamera_SDK_Stdcall" () As Long
Declare Function BUFUSB_StartFrameGrab Lib "BUF_USBCamera_SDK_Stdcall" (ByVal TotalFrames As Long) As Long
Declare Function BUFUSB_StopFrameGrab Lib "BUF_USBCamera_SDK_Stdcall" () As Long
Declare Function BUFUSB_SetResolution Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal Resolution As Long, ByVal Bin As Long, _
                                                                        ByVal BufferCnt As Long) As Long
Declare Function BUFUSB_SetCustomizedResolution Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal RowSize As Long, _
                                                                                  ByVal ColSize As Long, ByVal Bin As Long, ByVal BufferCnt As Long) As Long
Declare Function BUFUSB_SetExposureTime Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal exposureTime As Long) As Long
Declare Function BUFUSB_SetXYStart Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal XStart As Long, ByVal YStart As Long) As Long
Declare Function BUFUSB_SetGains Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal RedGain As Long, ByVal GreenGain As Long, _
                                                                   ByVal BlueGain As Long) As Long
Declare Function BUFUSB_SetGamma Lib "BUF_USBCamera_SDK_Stdcall" (ByVal Gamma As Long, ByVal Contrast As Long, ByVal Bright As Long, ByVal Sharp As Long) As Long
Declare Function BUFUSB_SetBWMode Lib "BUF_USBCamera_SDK_Stdcall" (ByVal BWMode As Long, ByVal H_Mirror As Long, ByVal V_Flip As Long) As Long
Declare Function BUFUSB_InstallFrameHooker Lib "BUF_USBCamera_SDK_Stdcall" (ByVal FrameType As Long, ByVal FrameHooker As Long) As Long
Declare Function BUFUSB_InstallUSBDeviceHooker Lib "BUF_USBCamera_SDK_Stdcall" (ByVal USBDeviceHooker As Long) As Long
Declare Function BUFUSB_SetGPIOConifg Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal ConfigByte As Byte) As Long
Declare Function BUFUSB_SetGPIOInOut Lib "BUF_USBCamera_SDK_Stdcall" (ByVal deviceID As Long, ByVal OutputByte As Byte, ByRef InputByteState As Byte) As Long
                                                                                  
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub Sleep Lib "kernel32" (ByVal ms&)
'

Sub MyFrameHooker(ByRef lpImageAttribute As TPROCESSEDDATAPROPERTY, ByRef Buffer As Byte)
    Dim MyStr As String

    Row = lpImageAttribute.Row / (lpImageAttribute.Bin + 1)
    Column = lpImageAttribute.Column / (lpImageAttribute.Bin + 1)
    
    FrameCount = FrameCount + 1
    ReDim ImageBuf(Column * Row)
       
    ' We copy the image data to ImageBuf()...Note that this callback is invoked in a higher
    ' priority thread... so this callback should not be blocked or do any GUI operations
    ' (Although this example does a simple GUI showing for demo purpose)...we recommend user to put the
    ' Image data and attributes in buffer and post a message to main thread for further
    ' processing.
    CopyMemory ImageBuf(0), Buffer, (Column * Row)
        
    MyStr = "Frame: " & FrameCount & "   Resolution: " & Column & "x" & Row
    Form2.Caption = "Video - " & MyStr
    
End Sub

Sub WriteImage()
      Dim i As Long, j As Long
      Dim CreateRow As String, Pixel As String
      Dim Mainpath As String, MIGHTEXDATA As String
      MIGHTEXDATA = "MightexImage.dat"
      ''ChDir (Mainpath)
      
      Open MIGHTEXDATA For Output As #45
                  For i = 0 To (Row - 1)
                        CreateRow = ""
                        For j = 0 To Column
                            If (ImageBuf(i * Column + j) < 16) Then
                            Pixel = "0" + Hex$(ImageBuf(i * Column + j))
                            Else
                            Pixel = Hex$(ImageBuf(i * Column + j))
                            End If
                            CreateRow = CreateRow + Pixel + " "
                        Next j
                        Print #45, CreateRow
                  Next i
      Close #45
End Sub

Sub DisplayImage()

    Dim pic As StdPicture, Pixels As Long, DarkPixels As Long, i As Long
    
    Pixels = Row * Column
    DarkPixels = 0
    
    'write camera data on top of template byte array

    If Form1.AreaCalc.Value And Form1.WhiteOut.Value Then
        For i = 1 To (Pixels - 1)
            If ImageBuf(i) < blacklevel Then
                DarkPixels = DarkPixels + 1     'count number of dark pixels
            Else
                ImageBuf(i) = 255               'set light pixels to white
            End If
            
            b(1311798 - i) = ImageBuf(i)
            Next i
    ElseIf Form1.AreaCalc.Value Then
        For i = 1 To (Pixels - 1)
            If ImageBuf(i) < blacklevel Then
                DarkPixels = DarkPixels + 1     'count number of dark pixels
            End If
            b(1311798 - i) = ImageBuf(i)
        Next i
    Else
        For i = 1 To (Pixels - 1)
            b(1311798 - i) = ImageBuf(i)
        Next i
    End If
    
    If Form1.AreaCalc.Value Then
        Form1.AreaDisplay.Text = Round(DarkPixels * PixelArea, 2) & " mm^2"
    End If
    
    
    Set pic = PictureFromByteStream(b)          'Create a StdPicture object (bitmap object) from the bytestream
    If pic Is Nothing Then
        MsgBox "Unable to load bitmap!"
        Exit Sub
    End If

    Set Form2.Picture1.Picture = pic

    Set pic = Nothing                           'Destroy pic when done

End Sub

