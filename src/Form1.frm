VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Area Determination"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   3000
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1440
      Top             =   0
   End
   Begin VB.CommandButton stop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin CWUIControlsLib.CWSlide BrightnessSlider 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
      _Version        =   524288
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   524288
      Slider_0        =   1
      ClassName_1     =   "CCWSlider"
      opts_1          =   2110
      C[0]_1          =   -2147483643
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
      Image_2         =   3
      ClassName_3     =   "CCWPictImage"
      opts_3          =   1280
      Rows_3          =   1
      Cols_3          =   1
      Pict_3          =   286
      F_3             =   -2147483633
      B_3             =   -2147483633
      ColorReplaceWith_3=   8421504
      ColorReplace_3  =   8421504
      Tolerance_3     =   2
      Animator_2      =   0
      Blinker_2       =   0
      BFImg_1         =   4
      ClassName_4     =   "CCWDrawObj"
      opts_4          =   62
      Image_4         =   5
      ClassName_5     =   "CCWPictImage"
      opts_5          =   1280
      Rows_5          =   1
      Cols_5          =   1
      Pict_5          =   286
      F_5             =   -2147483633
      B_5             =   -2147483633
      ColorReplaceWith_5=   8421504
      ColorReplace_5  =   8421504
      Tolerance_5     =   2
      Animator_4      =   0
      Blinker_4       =   0
      style_1         =   1
      Label_1         =   6
      ClassName_6     =   "CCWDrawObj"
      opts_6          =   62
      C[0]_6          =   -2147483640
      Image_6         =   7
      ClassName_7     =   "CCWTextImage"
      szText_7        =   "Brightness"
      font_7          =   0
      Animator_6      =   0
      Blinker_6       =   0
      Border_1        =   8
      ClassName_8     =   "CCWDrawObj"
      opts_8          =   60
      Image_8         =   9
      ClassName_9     =   "CCWPictImage"
      opts_9          =   1280
      Rows_9          =   1
      Cols_9          =   1
      Pict_9          =   25
      F_9             =   -2147483633
      B_9             =   -2147483633
      ColorReplaceWith_9=   8421504
      ColorReplace_9  =   8421504
      Tolerance_9     =   2
      Animator_8      =   0
      Blinker_8       =   0
      FillBound_1     =   10
      ClassName_10    =   "CCWGuiObject"
      opts_10         =   60
      FillTok_1       =   11
      ClassName_11    =   "CCWGuiObject"
      opts_11         =   62
      Axis_1          =   12
      ClassName_12    =   "CCWAxis"
      opts_12         =   575
      Name_12         =   "Axis"
      Orientation_12  =   1794
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   24576
      rMin_14         =   10
      rMax_14         =   174
      dMax_14         =   100
      discInterval_14 =   1
      Radial_12       =   0
      Enum_12         =   15
      ClassName_15    =   "CCWEnum"
      Editor_15       =   16
      ClassName_16    =   "CCWEnumArrayEditor"
      Owner_16        =   12
      Font_12         =   0
      tickopts_12     =   1678
      major_12        =   2
      minor_12        =   2
      Caption_12      =   17
      ClassName_17    =   "CCWDrawObj"
      opts_17         =   62
      C[0]_17         =   -2147483640
      Image_17        =   18
      ClassName_18    =   "CCWTextImage"
      font_18         =   0
      Animator_17     =   0
      Blinker_17      =   0
      DrawLst_1       =   19
      ClassName_19    =   "CDrawList"
      count_19        =   10
      list[10]_19     =   8
      list[9]_19      =   20
      ClassName_20    =   "CCWThumb"
      opts_20         =   63
      Name_20         =   "Pointer-1"
      C[2]_20         =   -2147483635
      Image_20        =   21
      ClassName_21    =   "CCWPictImage"
      opts_21         =   1280
      Rows_21         =   1
      Cols_21         =   1
      Pict_21         =   218
      F_21            =   -2147483633
      B_21            =   -2147483633
      ColorReplaceWith_21=   8421504
      ColorReplace_21 =   8421504
      Tolerance_21    =   2
      Animator_20     =   0
      Blinker_20      =   0
      style_20        =   1
      Value_20        =   50
      FillStyle_20    =   1
      Fill_20         =   22
      ClassName_22    =   "CCWDrawObj"
      opts_22         =   62
      C[0]_22         =   -2147483635
      Image_22        =   23
      ClassName_23    =   "CCWPictImage"
      opts_23         =   1280
      Rows_23         =   1
      Cols_23         =   1
      Pict_23         =   286
      F_23            =   -2147483635
      B_23            =   -2147483633
      ColorReplaceWith_23=   8421504
      ColorReplace_23 =   8421504
      Tolerance_23    =   2
      Animator_22     =   0
      Blinker_22      =   0
      list[8]_19      =   12
      list[7]_19      =   6
      list[6]_19      =   11
      list[5]_19      =   4
      list[4]_19      =   24
      ClassName_24    =   "CCWDrawObj"
      opts_24         =   60
      Image_24        =   25
      ClassName_25    =   "CCWPictImage"
      opts_25         =   1280
      Rows_25         =   1
      Cols_25         =   1
      Pict_25         =   94
      F_25            =   -2147483633
      B_25            =   -2147483633
      ColorReplaceWith_25=   8421504
      ColorReplace_25 =   8421504
      Tolerance_25    =   2
      Animator_24     =   0
      Blinker_24      =   0
      list[3]_19      =   26
      ClassName_26    =   "CCWDrawObj"
      opts_26         =   60
      Image_26        =   27
      ClassName_27    =   "CCWPictImage"
      opts_27         =   1280
      Rows_27         =   1
      Cols_27         =   1
      Pict_27         =   93
      F_27            =   -2147483633
      B_27            =   -2147483633
      ColorReplaceWith_27=   8421504
      ColorReplace_27 =   8421504
      Tolerance_27    =   2
      Animator_26     =   0
      Blinker_26      =   0
      list[2]_19      =   28
      ClassName_28    =   "CCWDrawObj"
      opts_28         =   62
      Image_28        =   29
      ClassName_29    =   "CCWPictImage"
      opts_29         =   1280
      Rows_29         =   1
      Cols_29         =   1
      Pict_29         =   7
      F_29            =   -2147483633
      B_29            =   -2147483633
      ColorReplaceWith_29=   8421504
      ColorReplace_29 =   8421504
      Tolerance_29    =   2
      Animator_28     =   0
      Blinker_28      =   0
      list[1]_19      =   2
      Ptrs_1          =   30
      ClassName_30    =   "CCWPointerArray"
      Array_30        =   1
      Editor_30       =   31
      ClassName_31    =   "CCWPointerArrayEditor"
      Owner_31        =   1
      Array[0]_30     =   20
      Bindings_1      =   32
      ClassName_32    =   "CCWBindingHolderArray"
      Editor_32       =   33
      ClassName_33    =   "CCWBindingHolderArrayEditor"
      Owner_33        =   1
      Stats_1         =   34
      ClassName_34    =   "CCWStats"
      doInc_1         =   24
      doDec_1         =   26
      doFrame_1       =   28
   End
   Begin VB.CommandButton start 
      Caption         =   "Start Video"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin CWUIControlsLib.CWSlide ContrastSlider 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2775
      _Version        =   524288
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   524288
      Slider_0        =   1
      ClassName_1     =   "CCWSlider"
      opts_1          =   2110
      C[0]_1          =   -2147483643
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
      Image_2         =   3
      ClassName_3     =   "CCWPictImage"
      opts_3          =   1280
      Rows_3          =   1
      Cols_3          =   1
      Pict_3          =   286
      F_3             =   -2147483633
      B_3             =   -2147483633
      ColorReplaceWith_3=   8421504
      ColorReplace_3  =   8421504
      Tolerance_3     =   2
      Animator_2      =   0
      Blinker_2       =   0
      BFImg_1         =   4
      ClassName_4     =   "CCWDrawObj"
      opts_4          =   62
      Image_4         =   5
      ClassName_5     =   "CCWPictImage"
      opts_5          =   1280
      Rows_5          =   1
      Cols_5          =   1
      Pict_5          =   286
      F_5             =   -2147483633
      B_5             =   -2147483633
      ColorReplaceWith_5=   8421504
      ColorReplace_5  =   8421504
      Tolerance_5     =   2
      Animator_4      =   0
      Blinker_4       =   0
      style_1         =   1
      Label_1         =   6
      ClassName_6     =   "CCWDrawObj"
      opts_6          =   62
      C[0]_6          =   -2147483640
      Image_6         =   7
      ClassName_7     =   "CCWTextImage"
      szText_7        =   "Contrast"
      font_7          =   0
      Animator_6      =   0
      Blinker_6       =   0
      Border_1        =   8
      ClassName_8     =   "CCWDrawObj"
      opts_8          =   60
      Image_8         =   9
      ClassName_9     =   "CCWPictImage"
      opts_9          =   1280
      Rows_9          =   1
      Cols_9          =   1
      Pict_9          =   25
      F_9             =   -2147483633
      B_9             =   -2147483633
      ColorReplaceWith_9=   8421504
      ColorReplace_9  =   8421504
      Tolerance_9     =   2
      Animator_8      =   0
      Blinker_8       =   0
      FillBound_1     =   10
      ClassName_10    =   "CCWGuiObject"
      opts_10         =   60
      FillTok_1       =   11
      ClassName_11    =   "CCWGuiObject"
      opts_11         =   62
      Axis_1          =   12
      ClassName_12    =   "CCWAxis"
      opts_12         =   575
      Name_12         =   "Axis"
      Orientation_12  =   1794
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   24576
      rMin_14         =   10
      rMax_14         =   174
      dMax_14         =   100
      discInterval_14 =   1
      Radial_12       =   0
      Enum_12         =   15
      ClassName_15    =   "CCWEnum"
      Editor_15       =   16
      ClassName_16    =   "CCWEnumArrayEditor"
      Owner_16        =   12
      Font_12         =   0
      tickopts_12     =   1678
      major_12        =   2
      minor_12        =   2
      Caption_12      =   17
      ClassName_17    =   "CCWDrawObj"
      opts_17         =   62
      C[0]_17         =   -2147483640
      Image_17        =   18
      ClassName_18    =   "CCWTextImage"
      font_18         =   0
      Animator_17     =   0
      Blinker_17      =   0
      DrawLst_1       =   19
      ClassName_19    =   "CDrawList"
      count_19        =   10
      list[10]_19     =   8
      list[9]_19      =   20
      ClassName_20    =   "CCWThumb"
      opts_20         =   63
      Name_20         =   "Pointer-1"
      C[2]_20         =   -2147483635
      Image_20        =   21
      ClassName_21    =   "CCWPictImage"
      opts_21         =   1280
      Rows_21         =   1
      Cols_21         =   1
      Pict_21         =   218
      F_21            =   -2147483633
      B_21            =   -2147483633
      ColorReplaceWith_21=   8421504
      ColorReplace_21 =   8421504
      Tolerance_21    =   2
      Animator_20     =   0
      Blinker_20      =   0
      style_20        =   1
      Value_20        =   50
      FillStyle_20    =   1
      Fill_20         =   22
      ClassName_22    =   "CCWDrawObj"
      opts_22         =   62
      C[0]_22         =   -2147483635
      Image_22        =   23
      ClassName_23    =   "CCWPictImage"
      opts_23         =   1280
      Rows_23         =   1
      Cols_23         =   1
      Pict_23         =   286
      F_23            =   -2147483635
      B_23            =   -2147483633
      ColorReplaceWith_23=   8421504
      ColorReplace_23 =   8421504
      Tolerance_23    =   2
      Animator_22     =   0
      Blinker_22      =   0
      list[8]_19      =   12
      list[7]_19      =   6
      list[6]_19      =   11
      list[5]_19      =   4
      list[4]_19      =   24
      ClassName_24    =   "CCWDrawObj"
      opts_24         =   60
      Image_24        =   25
      ClassName_25    =   "CCWPictImage"
      opts_25         =   1280
      Rows_25         =   1
      Cols_25         =   1
      Pict_25         =   94
      F_25            =   -2147483633
      B_25            =   -2147483633
      ColorReplaceWith_25=   8421504
      ColorReplace_25 =   8421504
      Tolerance_25    =   2
      Animator_24     =   0
      Blinker_24      =   0
      list[3]_19      =   26
      ClassName_26    =   "CCWDrawObj"
      opts_26         =   60
      Image_26        =   27
      ClassName_27    =   "CCWPictImage"
      opts_27         =   1280
      Rows_27         =   1
      Cols_27         =   1
      Pict_27         =   93
      F_27            =   -2147483633
      B_27            =   -2147483633
      ColorReplaceWith_27=   8421504
      ColorReplace_27 =   8421504
      Tolerance_27    =   2
      Animator_26     =   0
      Blinker_26      =   0
      list[2]_19      =   28
      ClassName_28    =   "CCWDrawObj"
      opts_28         =   62
      Image_28        =   29
      ClassName_29    =   "CCWPictImage"
      opts_29         =   1280
      Rows_29         =   1
      Cols_29         =   1
      Pict_29         =   7
      F_29            =   -2147483633
      B_29            =   -2147483633
      ColorReplaceWith_29=   8421504
      ColorReplace_29 =   8421504
      Tolerance_29    =   2
      Animator_28     =   0
      Blinker_28      =   0
      list[1]_19      =   2
      Ptrs_1          =   30
      ClassName_30    =   "CCWPointerArray"
      Array_30        =   1
      Editor_30       =   31
      ClassName_31    =   "CCWPointerArrayEditor"
      Owner_31        =   1
      Array[0]_30     =   20
      Bindings_1      =   32
      ClassName_32    =   "CCWBindingHolderArray"
      Editor_32       =   33
      ClassName_33    =   "CCWBindingHolderArrayEditor"
      Owner_33        =   1
      Stats_1         =   34
      ClassName_34    =   "CCWStats"
      doInc_1         =   24
      doDec_1         =   26
      doFrame_1       =   28
   End
   Begin CWUIControlsLib.CWSlide GammaSlider 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   2775
      _Version        =   524288
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   524288
      Slider_0        =   1
      ClassName_1     =   "CCWSlider"
      opts_1          =   2110
      C[0]_1          =   -2147483643
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
      Image_2         =   3
      ClassName_3     =   "CCWPictImage"
      opts_3          =   1280
      Rows_3          =   1
      Cols_3          =   1
      Pict_3          =   286
      F_3             =   -2147483633
      B_3             =   -2147483633
      ColorReplaceWith_3=   8421504
      ColorReplace_3  =   8421504
      Tolerance_3     =   2
      Animator_2      =   0
      Blinker_2       =   0
      BFImg_1         =   4
      ClassName_4     =   "CCWDrawObj"
      opts_4          =   62
      Image_4         =   5
      ClassName_5     =   "CCWPictImage"
      opts_5          =   1280
      Rows_5          =   1
      Cols_5          =   1
      Pict_5          =   286
      F_5             =   -2147483633
      B_5             =   -2147483633
      ColorReplaceWith_5=   8421504
      ColorReplace_5  =   8421504
      Tolerance_5     =   2
      Animator_4      =   0
      Blinker_4       =   0
      style_1         =   1
      Label_1         =   6
      ClassName_6     =   "CCWDrawObj"
      opts_6          =   62
      C[0]_6          =   -2147483640
      Image_6         =   7
      ClassName_7     =   "CCWTextImage"
      szText_7        =   "Gamma"
      font_7          =   0
      Animator_6      =   0
      Blinker_6       =   0
      Border_1        =   8
      ClassName_8     =   "CCWDrawObj"
      opts_8          =   60
      Image_8         =   9
      ClassName_9     =   "CCWPictImage"
      opts_9          =   1280
      Rows_9          =   1
      Cols_9          =   1
      Pict_9          =   25
      F_9             =   -2147483633
      B_9             =   -2147483633
      ColorReplaceWith_9=   8421504
      ColorReplace_9  =   8421504
      Tolerance_9     =   2
      Animator_8      =   0
      Blinker_8       =   0
      FillBound_1     =   10
      ClassName_10    =   "CCWGuiObject"
      opts_10         =   60
      FillTok_1       =   11
      ClassName_11    =   "CCWGuiObject"
      opts_11         =   62
      Axis_1          =   12
      ClassName_12    =   "CCWAxis"
      opts_12         =   575
      Name_12         =   "Axis"
      Orientation_12  =   1794
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   24576
      rMin_14         =   10
      rMax_14         =   174
      dMin_14         =   1
      dMax_14         =   20
      discInterval_14 =   1
      Radial_12       =   0
      Enum_12         =   15
      ClassName_15    =   "CCWEnum"
      Editor_15       =   16
      ClassName_16    =   "CCWEnumArrayEditor"
      Owner_16        =   12
      Font_12         =   0
      tickopts_12     =   1678
      major_12        =   2
      minor_12        =   2
      Caption_12      =   17
      ClassName_17    =   "CCWDrawObj"
      opts_17         =   62
      C[0]_17         =   -2147483640
      Image_17        =   18
      ClassName_18    =   "CCWTextImage"
      font_18         =   0
      Animator_17     =   0
      Blinker_17      =   0
      DrawLst_1       =   19
      ClassName_19    =   "CDrawList"
      count_19        =   10
      list[10]_19     =   8
      list[9]_19      =   20
      ClassName_20    =   "CCWThumb"
      opts_20         =   63
      Name_20         =   "Pointer-1"
      C[2]_20         =   -2147483635
      Image_20        =   21
      ClassName_21    =   "CCWPictImage"
      opts_21         =   1280
      Rows_21         =   1
      Cols_21         =   1
      Pict_21         =   218
      F_21            =   -2147483633
      B_21            =   -2147483633
      ColorReplaceWith_21=   8421504
      ColorReplace_21 =   8421504
      Tolerance_21    =   2
      Animator_20     =   0
      Blinker_20      =   0
      style_20        =   1
      Value_20        =   10
      FillStyle_20    =   1
      Fill_20         =   22
      ClassName_22    =   "CCWDrawObj"
      opts_22         =   62
      C[0]_22         =   -2147483635
      Image_22        =   23
      ClassName_23    =   "CCWPictImage"
      opts_23         =   1280
      Rows_23         =   1
      Cols_23         =   1
      Pict_23         =   286
      F_23            =   -2147483635
      B_23            =   -2147483633
      ColorReplaceWith_23=   8421504
      ColorReplace_23 =   8421504
      Tolerance_23    =   2
      Animator_22     =   0
      Blinker_22      =   0
      list[8]_19      =   12
      list[7]_19      =   6
      list[6]_19      =   11
      list[5]_19      =   4
      list[4]_19      =   24
      ClassName_24    =   "CCWDrawObj"
      opts_24         =   60
      Image_24        =   25
      ClassName_25    =   "CCWPictImage"
      opts_25         =   1280
      Rows_25         =   1
      Cols_25         =   1
      Pict_25         =   94
      F_25            =   -2147483633
      B_25            =   -2147483633
      ColorReplaceWith_25=   8421504
      ColorReplace_25 =   8421504
      Tolerance_25    =   2
      Animator_24     =   0
      Blinker_24      =   0
      list[3]_19      =   26
      ClassName_26    =   "CCWDrawObj"
      opts_26         =   60
      Image_26        =   27
      ClassName_27    =   "CCWPictImage"
      opts_27         =   1280
      Rows_27         =   1
      Cols_27         =   1
      Pict_27         =   93
      F_27            =   -2147483633
      B_27            =   -2147483633
      ColorReplaceWith_27=   8421504
      ColorReplace_27 =   8421504
      Tolerance_27    =   2
      Animator_26     =   0
      Blinker_26      =   0
      list[2]_19      =   28
      ClassName_28    =   "CCWDrawObj"
      opts_28         =   62
      Image_28        =   29
      ClassName_29    =   "CCWPictImage"
      opts_29         =   1280
      Rows_29         =   1
      Cols_29         =   1
      Pict_29         =   7
      F_29            =   -2147483633
      B_29            =   -2147483633
      ColorReplaceWith_29=   8421504
      ColorReplace_29 =   8421504
      Tolerance_29    =   2
      Animator_28     =   0
      Blinker_28      =   0
      list[1]_19      =   2
      Ptrs_1          =   30
      ClassName_30    =   "CCWPointerArray"
      Array_30        =   1
      Editor_30       =   31
      ClassName_31    =   "CCWPointerArrayEditor"
      Owner_31        =   1
      Array[0]_30     =   20
      Bindings_1      =   32
      ClassName_32    =   "CCWBindingHolderArray"
      Editor_32       =   33
      ClassName_33    =   "CCWBindingHolderArrayEditor"
      Owner_33        =   1
      Stats_1         =   34
      ClassName_34    =   "CCWStats"
      doInc_1         =   24
      doDec_1         =   26
      doFrame_1       =   28
   End
   Begin CWUIControlsLib.CWSlide GainSlider 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   2775
      _Version        =   524288
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   524288
      Slider_0        =   1
      ClassName_1     =   "CCWSlider"
      opts_1          =   2110
      C[0]_1          =   -2147483643
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
      Image_2         =   3
      ClassName_3     =   "CCWPictImage"
      opts_3          =   1280
      Rows_3          =   1
      Cols_3          =   1
      Pict_3          =   286
      F_3             =   -2147483633
      B_3             =   -2147483633
      ColorReplaceWith_3=   8421504
      ColorReplace_3  =   8421504
      Tolerance_3     =   2
      Animator_2      =   0
      Blinker_2       =   0
      BFImg_1         =   4
      ClassName_4     =   "CCWDrawObj"
      opts_4          =   62
      Image_4         =   5
      ClassName_5     =   "CCWPictImage"
      opts_5          =   1280
      Rows_5          =   1
      Cols_5          =   1
      Pict_5          =   286
      F_5             =   -2147483633
      B_5             =   -2147483633
      ColorReplaceWith_5=   8421504
      ColorReplace_5  =   8421504
      Tolerance_5     =   2
      Animator_4      =   0
      Blinker_4       =   0
      style_1         =   1
      Label_1         =   6
      ClassName_6     =   "CCWDrawObj"
      opts_6          =   62
      C[0]_6          =   -2147483640
      Image_6         =   7
      ClassName_7     =   "CCWTextImage"
      szText_7        =   "Gain"
      font_7          =   0
      Animator_6      =   0
      Blinker_6       =   0
      Border_1        =   8
      ClassName_8     =   "CCWDrawObj"
      opts_8          =   60
      Image_8         =   9
      ClassName_9     =   "CCWPictImage"
      opts_9          =   1280
      Rows_9          =   1
      Cols_9          =   1
      Pict_9          =   25
      F_9             =   -2147483633
      B_9             =   -2147483633
      ColorReplaceWith_9=   8421504
      ColorReplace_9  =   8421504
      Tolerance_9     =   2
      Animator_8      =   0
      Blinker_8       =   0
      FillBound_1     =   10
      ClassName_10    =   "CCWGuiObject"
      opts_10         =   60
      FillTok_1       =   11
      ClassName_11    =   "CCWGuiObject"
      opts_11         =   62
      Axis_1          =   12
      ClassName_12    =   "CCWAxis"
      opts_12         =   575
      Name_12         =   "Axis"
      Orientation_12  =   1794
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   24576
      rMin_14         =   10
      rMax_14         =   174
      dMin_14         =   1
      dMax_14         =   64
      discInterval_14 =   1
      Radial_12       =   0
      Enum_12         =   15
      ClassName_15    =   "CCWEnum"
      Editor_15       =   16
      ClassName_16    =   "CCWEnumArrayEditor"
      Owner_16        =   12
      Font_12         =   0
      tickopts_12     =   1678
      major_12        =   2
      minor_12        =   2
      Caption_12      =   17
      ClassName_17    =   "CCWDrawObj"
      opts_17         =   62
      C[0]_17         =   -2147483640
      Image_17        =   18
      ClassName_18    =   "CCWTextImage"
      font_18         =   0
      Animator_17     =   0
      Blinker_17      =   0
      DrawLst_1       =   19
      ClassName_19    =   "CDrawList"
      count_19        =   10
      list[10]_19     =   8
      list[9]_19      =   20
      ClassName_20    =   "CCWThumb"
      opts_20         =   63
      Name_20         =   "Pointer-1"
      C[2]_20         =   -2147483635
      Image_20        =   21
      ClassName_21    =   "CCWPictImage"
      opts_21         =   1280
      Rows_21         =   1
      Cols_21         =   1
      Pict_21         =   218
      F_21            =   -2147483633
      B_21            =   -2147483633
      ColorReplaceWith_21=   8421504
      ColorReplace_21 =   8421504
      Tolerance_21    =   2
      Animator_20     =   0
      Blinker_20      =   0
      style_20        =   1
      Value_20        =   32
      FillStyle_20    =   1
      Fill_20         =   22
      ClassName_22    =   "CCWDrawObj"
      opts_22         =   62
      C[0]_22         =   -2147483635
      Image_22        =   23
      ClassName_23    =   "CCWPictImage"
      opts_23         =   1280
      Rows_23         =   1
      Cols_23         =   1
      Pict_23         =   286
      F_23            =   -2147483635
      B_23            =   -2147483633
      ColorReplaceWith_23=   8421504
      ColorReplace_23 =   8421504
      Tolerance_23    =   2
      Animator_22     =   0
      Blinker_22      =   0
      list[8]_19      =   12
      list[7]_19      =   6
      list[6]_19      =   11
      list[5]_19      =   4
      list[4]_19      =   24
      ClassName_24    =   "CCWDrawObj"
      opts_24         =   60
      Image_24        =   25
      ClassName_25    =   "CCWPictImage"
      opts_25         =   1280
      Rows_25         =   1
      Cols_25         =   1
      Pict_25         =   94
      F_25            =   -2147483633
      B_25            =   -2147483633
      ColorReplaceWith_25=   8421504
      ColorReplace_25 =   8421504
      Tolerance_25    =   2
      Animator_24     =   0
      Blinker_24      =   0
      list[3]_19      =   26
      ClassName_26    =   "CCWDrawObj"
      opts_26         =   60
      Image_26        =   27
      ClassName_27    =   "CCWPictImage"
      opts_27         =   1280
      Rows_27         =   1
      Cols_27         =   1
      Pict_27         =   93
      F_27            =   -2147483633
      B_27            =   -2147483633
      ColorReplaceWith_27=   8421504
      ColorReplace_27 =   8421504
      Tolerance_27    =   2
      Animator_26     =   0
      Blinker_26      =   0
      list[2]_19      =   28
      ClassName_28    =   "CCWDrawObj"
      opts_28         =   62
      Image_28        =   29
      ClassName_29    =   "CCWPictImage"
      opts_29         =   1280
      Rows_29         =   1
      Cols_29         =   1
      Pict_29         =   7
      F_29            =   -2147483633
      B_29            =   -2147483633
      ColorReplaceWith_29=   8421504
      ColorReplace_29 =   8421504
      Tolerance_29    =   2
      Animator_28     =   0
      Blinker_28      =   0
      list[1]_19      =   2
      Ptrs_1          =   30
      ClassName_30    =   "CCWPointerArray"
      Array_30        =   1
      Editor_30       =   31
      ClassName_31    =   "CCWPointerArrayEditor"
      Owner_31        =   1
      Array[0]_30     =   20
      Bindings_1      =   32
      ClassName_32    =   "CCWBindingHolderArray"
      Editor_32       =   33
      ClassName_33    =   "CCWBindingHolderArrayEditor"
      Owner_33        =   1
      Stats_1         =   34
      ClassName_34    =   "CCWStats"
      doInc_1         =   24
      doDec_1         =   26
      doFrame_1       =   28
   End
   Begin CWUIControlsLib.CWSlide ExposureSlider 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   2775
      _Version        =   524288
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   524288
      Slider_0        =   1
      ClassName_1     =   "CCWSlider"
      opts_1          =   2110
      C[0]_1          =   -2147483643
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
      Image_2         =   3
      ClassName_3     =   "CCWPictImage"
      opts_3          =   1280
      Rows_3          =   1
      Cols_3          =   1
      Pict_3          =   286
      F_3             =   -2147483633
      B_3             =   -2147483633
      ColorReplaceWith_3=   8421504
      ColorReplace_3  =   8421504
      Tolerance_3     =   2
      Animator_2      =   0
      Blinker_2       =   0
      BFImg_1         =   4
      ClassName_4     =   "CCWDrawObj"
      opts_4          =   62
      Image_4         =   5
      ClassName_5     =   "CCWPictImage"
      opts_5          =   1280
      Rows_5          =   1
      Cols_5          =   1
      Pict_5          =   286
      F_5             =   -2147483633
      B_5             =   -2147483633
      ColorReplaceWith_5=   8421504
      ColorReplace_5  =   8421504
      Tolerance_5     =   2
      Animator_4      =   0
      Blinker_4       =   0
      style_1         =   1
      Label_1         =   6
      ClassName_6     =   "CCWDrawObj"
      opts_6          =   62
      C[0]_6          =   -2147483640
      Image_6         =   7
      ClassName_7     =   "CCWTextImage"
      szText_7        =   "Exposure Time"
      font_7          =   0
      Animator_6      =   0
      Blinker_6       =   0
      Border_1        =   8
      ClassName_8     =   "CCWDrawObj"
      opts_8          =   60
      Image_8         =   9
      ClassName_9     =   "CCWPictImage"
      opts_9          =   1280
      Rows_9          =   1
      Cols_9          =   1
      Pict_9          =   25
      F_9             =   -2147483633
      B_9             =   -2147483633
      ColorReplaceWith_9=   8421504
      ColorReplace_9  =   8421504
      Tolerance_9     =   2
      Animator_8      =   0
      Blinker_8       =   0
      FillBound_1     =   10
      ClassName_10    =   "CCWGuiObject"
      opts_10         =   60
      FillTok_1       =   11
      ClassName_11    =   "CCWGuiObject"
      opts_11         =   62
      Axis_1          =   12
      ClassName_12    =   "CCWAxis"
      opts_12         =   575
      Name_12         =   "Axis"
      Orientation_12  =   1794
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   24576
      rMin_14         =   10
      rMax_14         =   174
      dMin_14         =   50
      dMax_14         =   750
      discInterval_14 =   1
      Radial_12       =   0
      Enum_12         =   15
      ClassName_15    =   "CCWEnum"
      Editor_15       =   16
      ClassName_16    =   "CCWEnumArrayEditor"
      Owner_16        =   12
      Font_12         =   0
      tickopts_12     =   1678
      major_12        =   2
      minor_12        =   2
      Caption_12      =   17
      ClassName_17    =   "CCWDrawObj"
      opts_17         =   62
      C[0]_17         =   -2147483640
      Image_17        =   18
      ClassName_18    =   "CCWTextImage"
      font_18         =   0
      Animator_17     =   0
      Blinker_17      =   0
      DrawLst_1       =   19
      ClassName_19    =   "CDrawList"
      count_19        =   10
      list[10]_19     =   8
      list[9]_19      =   20
      ClassName_20    =   "CCWThumb"
      opts_20         =   63
      Name_20         =   "Pointer-1"
      C[2]_20         =   -2147483635
      Image_20        =   21
      ClassName_21    =   "CCWPictImage"
      opts_21         =   1280
      Rows_21         =   1
      Cols_21         =   1
      Pict_21         =   218
      F_21            =   -2147483633
      B_21            =   -2147483633
      ColorReplaceWith_21=   8421504
      ColorReplace_21 =   8421504
      Tolerance_21    =   2
      Animator_20     =   0
      Blinker_20      =   0
      style_20        =   1
      Value_20        =   375
      FillStyle_20    =   1
      Fill_20         =   22
      ClassName_22    =   "CCWDrawObj"
      opts_22         =   62
      C[0]_22         =   -2147483635
      Image_22        =   23
      ClassName_23    =   "CCWPictImage"
      opts_23         =   1280
      Rows_23         =   1
      Cols_23         =   1
      Pict_23         =   286
      F_23            =   -2147483635
      B_23            =   -2147483633
      ColorReplaceWith_23=   8421504
      ColorReplace_23 =   8421504
      Tolerance_23    =   2
      Animator_22     =   0
      Blinker_22      =   0
      list[8]_19      =   12
      list[7]_19      =   6
      list[6]_19      =   11
      list[5]_19      =   4
      list[4]_19      =   24
      ClassName_24    =   "CCWDrawObj"
      opts_24         =   60
      Image_24        =   25
      ClassName_25    =   "CCWPictImage"
      opts_25         =   1280
      Rows_25         =   1
      Cols_25         =   1
      Pict_25         =   94
      F_25            =   -2147483633
      B_25            =   -2147483633
      ColorReplaceWith_25=   8421504
      ColorReplace_25 =   8421504
      Tolerance_25    =   2
      Animator_24     =   0
      Blinker_24      =   0
      list[3]_19      =   26
      ClassName_26    =   "CCWDrawObj"
      opts_26         =   60
      Image_26        =   27
      ClassName_27    =   "CCWPictImage"
      opts_27         =   1280
      Rows_27         =   1
      Cols_27         =   1
      Pict_27         =   93
      F_27            =   -2147483633
      B_27            =   -2147483633
      ColorReplaceWith_27=   8421504
      ColorReplace_27 =   8421504
      Tolerance_27    =   2
      Animator_26     =   0
      Blinker_26      =   0
      list[2]_19      =   28
      ClassName_28    =   "CCWDrawObj"
      opts_28         =   62
      Image_28        =   29
      ClassName_29    =   "CCWPictImage"
      opts_29         =   1280
      Rows_29         =   1
      Cols_29         =   1
      Pict_29         =   7
      F_29            =   -2147483633
      B_29            =   -2147483633
      ColorReplaceWith_29=   8421504
      ColorReplace_29 =   8421504
      Tolerance_29    =   2
      Animator_28     =   0
      Blinker_28      =   0
      list[1]_19      =   2
      Ptrs_1          =   30
      ClassName_30    =   "CCWPointerArray"
      Array_30        =   1
      Editor_30       =   31
      ClassName_31    =   "CCWPointerArrayEditor"
      Owner_31        =   1
      Array[0]_30     =   20
      Bindings_1      =   32
      ClassName_32    =   "CCWBindingHolderArray"
      Editor_32       =   33
      ClassName_33    =   "CCWBindingHolderArrayEditor"
      Owner_33        =   1
      Stats_1         =   34
      ClassName_34    =   "CCWStats"
      doInc_1         =   24
      doDec_1         =   26
      doFrame_1       =   28
   End
   Begin CWUIControlsLib.CWSlide BlackSlider 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   2775
      _Version        =   524288
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   524288
      Slider_0        =   1
      ClassName_1     =   "CCWSlider"
      opts_1          =   2110
      C[0]_1          =   -2147483643
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
      Image_2         =   3
      ClassName_3     =   "CCWPictImage"
      opts_3          =   1280
      Rows_3          =   1
      Cols_3          =   1
      Pict_3          =   286
      F_3             =   -2147483633
      B_3             =   -2147483633
      ColorReplaceWith_3=   8421504
      ColorReplace_3  =   8421504
      Tolerance_3     =   2
      Animator_2      =   0
      Blinker_2       =   0
      BFImg_1         =   4
      ClassName_4     =   "CCWDrawObj"
      opts_4          =   62
      Image_4         =   5
      ClassName_5     =   "CCWPictImage"
      opts_5          =   1280
      Rows_5          =   1
      Cols_5          =   1
      Pict_5          =   286
      F_5             =   -2147483633
      B_5             =   -2147483633
      ColorReplaceWith_5=   8421504
      ColorReplace_5  =   8421504
      Tolerance_5     =   2
      Animator_4      =   0
      Blinker_4       =   0
      style_1         =   1
      Label_1         =   6
      ClassName_6     =   "CCWDrawObj"
      opts_6          =   62
      C[0]_6          =   -2147483640
      Image_6         =   7
      ClassName_7     =   "CCWTextImage"
      szText_7        =   "Black Cutoff"
      font_7          =   0
      Animator_6      =   0
      Blinker_6       =   0
      Border_1        =   8
      ClassName_8     =   "CCWDrawObj"
      opts_8          =   60
      Image_8         =   9
      ClassName_9     =   "CCWPictImage"
      opts_9          =   1280
      Rows_9          =   1
      Cols_9          =   1
      Pict_9          =   25
      F_9             =   -2147483633
      B_9             =   -2147483633
      ColorReplaceWith_9=   8421504
      ColorReplace_9  =   8421504
      Tolerance_9     =   2
      Animator_8      =   0
      Blinker_8       =   0
      FillBound_1     =   10
      ClassName_10    =   "CCWGuiObject"
      opts_10         =   60
      FillTok_1       =   11
      ClassName_11    =   "CCWGuiObject"
      opts_11         =   62
      Axis_1          =   12
      ClassName_12    =   "CCWAxis"
      opts_12         =   575
      Name_12         =   "Axis"
      Orientation_12  =   1794
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   24576
      rMin_14         =   10
      rMax_14         =   174
      dMax_14         =   255
      discInterval_14 =   1
      Radial_12       =   0
      Enum_12         =   15
      ClassName_15    =   "CCWEnum"
      Editor_15       =   16
      ClassName_16    =   "CCWEnumArrayEditor"
      Owner_16        =   12
      Font_12         =   0
      tickopts_12     =   1678
      major_12        =   2
      minor_12        =   2
      Caption_12      =   17
      ClassName_17    =   "CCWDrawObj"
      opts_17         =   62
      C[0]_17         =   -2147483640
      Image_17        =   18
      ClassName_18    =   "CCWTextImage"
      font_18         =   0
      Animator_17     =   0
      Blinker_17      =   0
      DrawLst_1       =   19
      ClassName_19    =   "CDrawList"
      count_19        =   10
      list[10]_19     =   8
      list[9]_19      =   20
      ClassName_20    =   "CCWThumb"
      opts_20         =   63
      Name_20         =   "Pointer-1"
      C[2]_20         =   0
      Image_20        =   21
      ClassName_21    =   "CCWPictImage"
      opts_21         =   1280
      Rows_21         =   1
      Cols_21         =   1
      Pict_21         =   218
      F_21            =   -2147483633
      B_21            =   -2147483633
      ColorReplaceWith_21=   8421504
      ColorReplace_21 =   8421504
      Tolerance_21    =   2
      Animator_20     =   0
      Blinker_20      =   0
      style_20        =   1
      Value_20        =   128
      FillStyle_20    =   1
      Fill_20         =   22
      ClassName_22    =   "CCWDrawObj"
      opts_22         =   62
      C[0]_22         =   0
      Image_22        =   23
      ClassName_23    =   "CCWPictImage"
      opts_23         =   1280
      Rows_23         =   1
      Cols_23         =   1
      Pict_23         =   286
      F_23            =   0
      B_23            =   -2147483633
      ColorReplaceWith_23=   8421504
      ColorReplace_23 =   8421504
      Tolerance_23    =   2
      Animator_22     =   0
      Blinker_22      =   0
      list[8]_19      =   12
      list[7]_19      =   6
      list[6]_19      =   11
      list[5]_19      =   4
      list[4]_19      =   24
      ClassName_24    =   "CCWDrawObj"
      opts_24         =   62
      Image_24        =   25
      ClassName_25    =   "CCWPictImage"
      opts_25         =   1280
      Rows_25         =   1
      Cols_25         =   1
      Pict_25         =   94
      F_25            =   -2147483633
      B_25            =   -2147483633
      ColorReplaceWith_25=   8421504
      ColorReplace_25 =   8421504
      Tolerance_25    =   2
      Animator_24     =   0
      Blinker_24      =   0
      list[3]_19      =   26
      ClassName_26    =   "CCWDrawObj"
      opts_26         =   62
      Image_26        =   27
      ClassName_27    =   "CCWPictImage"
      opts_27         =   1280
      Rows_27         =   1
      Cols_27         =   1
      Pict_27         =   93
      F_27            =   -2147483633
      B_27            =   -2147483633
      ColorReplaceWith_27=   8421504
      ColorReplace_27 =   8421504
      Tolerance_27    =   2
      Animator_26     =   0
      Blinker_26      =   0
      list[2]_19      =   28
      ClassName_28    =   "CCWDrawObj"
      opts_28         =   62
      Image_28        =   29
      ClassName_29    =   "CCWPictImage"
      opts_29         =   1280
      Rows_29         =   1
      Cols_29         =   1
      Pict_29         =   7
      F_29            =   -2147483633
      B_29            =   -2147483633
      ColorReplaceWith_29=   8421504
      ColorReplace_29 =   8421504
      Tolerance_29    =   2
      Animator_28     =   0
      Blinker_28      =   0
      list[1]_19      =   2
      Ptrs_1          =   30
      ClassName_30    =   "CCWPointerArray"
      Array_30        =   1
      Editor_30       =   31
      ClassName_31    =   "CCWPointerArrayEditor"
      Owner_31        =   1
      Array[0]_30     =   20
      Bindings_1      =   32
      ClassName_32    =   "CCWBindingHolderArray"
      Editor_32       =   33
      ClassName_33    =   "CCWBindingHolderArrayEditor"
      Owner_33        =   1
      Stats_1         =   34
      ClassName_34    =   "CCWStats"
      doInc_1         =   24
      doDec_1         =   26
      doFrame_1       =   28
   End
   Begin VB.Frame Frame2 
      Caption         =   "Area Determination"
      Height          =   3255
      Left            =   0
      TabIndex        =   10
      Top             =   6000
      Width           =   3015
      Begin VB.CommandButton Calibrate 
         Caption         =   "Recalibrate"
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CheckBox WhiteOut 
         Caption         =   "White out"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox AreaCalc 
         Caption         =   "Calculate area"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox AreaDisplay 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label CalLabel 
         Caption         =   "Current Calibration: __ pixels/mm"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   2775
      End
   End
   Begin CWUIControlsLib.CWSlide SharpnessSlider 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   2775
      _Version        =   524288
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   524288
      Slider_0        =   1
      ClassName_1     =   "CCWSlider"
      opts_1          =   2110
      C[0]_1          =   -2147483643
      BGImg_1         =   2
      ClassName_2     =   "CCWDrawObj"
      opts_2          =   62
      Image_2         =   3
      ClassName_3     =   "CCWPictImage"
      opts_3          =   1280
      Rows_3          =   1
      Cols_3          =   1
      Pict_3          =   286
      F_3             =   -2147483633
      B_3             =   -2147483633
      ColorReplaceWith_3=   8421504
      ColorReplace_3  =   8421504
      Tolerance_3     =   2
      Animator_2      =   0
      Blinker_2       =   0
      BFImg_1         =   4
      ClassName_4     =   "CCWDrawObj"
      opts_4          =   62
      Image_4         =   5
      ClassName_5     =   "CCWPictImage"
      opts_5          =   1280
      Rows_5          =   1
      Cols_5          =   1
      Pict_5          =   286
      F_5             =   -2147483633
      B_5             =   -2147483633
      ColorReplaceWith_5=   8421504
      ColorReplace_5  =   8421504
      Tolerance_5     =   2
      Animator_4      =   0
      Blinker_4       =   0
      style_1         =   1
      Label_1         =   6
      ClassName_6     =   "CCWDrawObj"
      opts_6          =   62
      C[0]_6          =   -2147483640
      Image_6         =   7
      ClassName_7     =   "CCWTextImage"
      szText_7        =   "Sharpness"
      font_7          =   0
      Animator_6      =   0
      Blinker_6       =   0
      Border_1        =   8
      ClassName_8     =   "CCWDrawObj"
      opts_8          =   60
      Image_8         =   9
      ClassName_9     =   "CCWPictImage"
      opts_9          =   1280
      Rows_9          =   1
      Cols_9          =   1
      Pict_9          =   25
      F_9             =   -2147483633
      B_9             =   -2147483633
      ColorReplaceWith_9=   8421504
      ColorReplace_9  =   8421504
      Tolerance_9     =   2
      Animator_8      =   0
      Blinker_8       =   0
      FillBound_1     =   10
      ClassName_10    =   "CCWGuiObject"
      opts_10         =   60
      FillTok_1       =   11
      ClassName_11    =   "CCWGuiObject"
      opts_11         =   62
      Axis_1          =   12
      ClassName_12    =   "CCWAxis"
      opts_12         =   575
      Name_12         =   "Axis"
      Orientation_12  =   1794
      format_12       =   13
      ClassName_13    =   "CCWFormat"
      Scale_12        =   14
      ClassName_14    =   "CCWScale"
      opts_14         =   24576
      rMin_14         =   10
      rMax_14         =   174
      dMax_14         =   3
      discInterval_14 =   1
      Radial_12       =   0
      Enum_12         =   15
      ClassName_15    =   "CCWEnum"
      Editor_15       =   16
      ClassName_16    =   "CCWEnumArrayEditor"
      Owner_16        =   12
      Font_12         =   0
      tickopts_12     =   1678
      major_12        =   3
      Caption_12      =   17
      ClassName_17    =   "CCWDrawObj"
      opts_17         =   62
      C[0]_17         =   -2147483640
      Image_17        =   18
      ClassName_18    =   "CCWTextImage"
      font_18         =   0
      Animator_17     =   0
      Blinker_17      =   0
      DrawLst_1       =   19
      ClassName_19    =   "CDrawList"
      count_19        =   10
      list[10]_19     =   8
      list[9]_19      =   20
      ClassName_20    =   "CCWThumb"
      opts_20         =   63
      Name_20         =   "Pointer-1"
      C[2]_20         =   -2147483635
      Image_20        =   21
      ClassName_21    =   "CCWPictImage"
      opts_21         =   1280
      Rows_21         =   1
      Cols_21         =   1
      Pict_21         =   218
      F_21            =   -2147483633
      B_21            =   -2147483633
      ColorReplaceWith_21=   8421504
      ColorReplace_21 =   8421504
      Tolerance_21    =   2
      Animator_20     =   0
      Blinker_20      =   0
      style_20        =   1
      Value_20        =   3
      FillStyle_20    =   1
      Fill_20         =   22
      ClassName_22    =   "CCWDrawObj"
      opts_22         =   62
      C[0]_22         =   -2147483635
      Image_22        =   23
      ClassName_23    =   "CCWPictImage"
      opts_23         =   1280
      Rows_23         =   1
      Cols_23         =   1
      Pict_23         =   286
      F_23            =   -2147483635
      B_23            =   -2147483633
      ColorReplaceWith_23=   8421504
      ColorReplace_23 =   8421504
      Tolerance_23    =   2
      Animator_22     =   0
      Blinker_22      =   0
      list[8]_19      =   12
      list[7]_19      =   6
      list[6]_19      =   11
      list[5]_19      =   4
      list[4]_19      =   24
      ClassName_24    =   "CCWDrawObj"
      opts_24         =   60
      Image_24        =   25
      ClassName_25    =   "CCWPictImage"
      opts_25         =   1280
      Rows_25         =   1
      Cols_25         =   1
      Pict_25         =   94
      F_25            =   -2147483633
      B_25            =   -2147483633
      ColorReplaceWith_25=   8421504
      ColorReplace_25 =   8421504
      Tolerance_25    =   2
      Animator_24     =   0
      Blinker_24      =   0
      list[3]_19      =   26
      ClassName_26    =   "CCWDrawObj"
      opts_26         =   60
      Image_26        =   27
      ClassName_27    =   "CCWPictImage"
      opts_27         =   1280
      Rows_27         =   1
      Cols_27         =   1
      Pict_27         =   93
      F_27            =   -2147483633
      B_27            =   -2147483633
      ColorReplaceWith_27=   8421504
      ColorReplace_27 =   8421504
      Tolerance_27    =   2
      Animator_26     =   0
      Blinker_26      =   0
      list[2]_19      =   28
      ClassName_28    =   "CCWDrawObj"
      opts_28         =   62
      Image_28        =   29
      ClassName_29    =   "CCWPictImage"
      opts_29         =   1280
      Rows_29         =   1
      Cols_29         =   1
      Pict_29         =   7
      F_29            =   -2147483633
      B_29            =   -2147483633
      ColorReplaceWith_29=   8421504
      ColorReplace_29 =   8421504
      Tolerance_29    =   2
      Animator_28     =   0
      Blinker_28      =   0
      list[1]_19      =   2
      Ptrs_1          =   30
      ClassName_30    =   "CCWPointerArray"
      Array_30        =   1
      Editor_30       =   31
      ClassName_31    =   "CCWPointerArrayEditor"
      Owner_31        =   1
      Array[0]_30     =   20
      Bindings_1      =   32
      ClassName_32    =   "CCWBindingHolderArray"
      Editor_32       =   33
      ClassName_33    =   "CCWBindingHolderArrayEditor"
      Owner_33        =   1
      Stats_1         =   34
      ClassName_34    =   "CCWStats"
      doInc_1         =   24
      doDec_1         =   26
      doFrame_1       =   28
   End
   Begin VB.Frame Frame1 
      Caption         =   "Camera Settings"
      Height          =   4695
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   3015
      Begin VB.CommandButton SetDefault 
         Caption         =   "Set as Default"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton LoadDefault 
         Caption         =   "Load Default"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   3960
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "T. Hennen"
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label ModuleNoLabel 
      Caption         =   "Module No.:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Gamma As Long, Contrast As Long, Bright As Long, Sharp As Long, Exposure As Long, Gain As Long
Dim video As Boolean
Const MIGHTEXPARAMETERS As String = "MightexParameters.txt"


Private Sub Calibrate_Click()
LeftPixelX = 2000
RightPixelX = 2000
AreaCalc.Value = False
Form3.Show
End Sub

Private Sub Form_Load()
    
    'Load a picture into a byte array.  The program reuses the header information but overwrites the pixel data
    b = LoadFile(App.Path & "\Template.bmp")
    
    'Get calibration from file
    Open CALIBRATIONFILE For Input As #42
        Input #42, PixelsPerMM
    Close #42
    Form1.CalLabel.Caption = "Current Calibration: " & PixelsPerMM & " pixels/mm"
    PixelArea = 1 / (PixelsPerMM * PixelsPerMM)
End Sub
Sub CameraSettingsFromFile()
      Dim sFo As New FileSystemObject
      Dim Dummy As String
      
      '' read parameters from file
      
      If sFo.FileExists(MIGHTEXPARAMETERS) Then
            Open MIGHTEXPARAMETERS For Input As #56
                  Input #56, Dummy, Exposure
                  Input #56, Dummy, Gain
                  Input #56, Dummy, Gamma
                  Input #56, Dummy, Contrast
                  Input #56, Dummy, Bright
                  Input #56, Dummy, Sharp
            Close #56
            
                'set sliders to reflect settings
                BrightnessSlider.Value = Bright
                BrightnessSlider.Caption = "Brightness (" & Bright & ")"
    
                ContrastSlider.Value = Contrast
                ContrastSlider.Caption = "Contrast (" & Contrast & ")"
    
                GammaSlider.Value = Gamma
                GammaSlider.Caption = "Gamma (" & Gamma & ")"
    
                ExposureSlider.Value = Exposure
                ExposureSlider.Caption = "Exposure Time(" & Exposure & " ms)"
    
                GainSlider.Value = Gain
                GainSlider.Caption = "Gain (" & Gain & ")"
                
                SharpnessSlider.Value = Sharp
                SharpnessSlider.Caption = "Sharpness (" & Sharp & ")"
                
                'update camera
                BUFUSB_SetExposureTime 1, Exposure / 0.05
                BUFUSB_SetGains 1, Gain, Gain, Gain
                Call BUFUSB_SetGamma(Gamma, Contrast, Bright, Sharp)
      Else
            MsgBox ("parameter file for camera does not exist")
      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    video = False                               'stops frame refresh loop
    BUFUSB_StopFrameGrab
    BUFUSB_InstallFrameHooker BMPDATA_IMAGE, 0
    BUFUSB_StopCameraEngine
    BUFUSB_UnInitDevice
    Unload Form2
    Unload Form3
End Sub


Private Sub LoadDefault_Click()
    CameraSettingsFromFile
End Sub
Private Sub SetDefault_Click()
            Open MIGHTEXPARAMETERS For Output As #55
                  Print #55, "Exposure Time ms (1-750),"; Exposure
                  Print #55, "Gain (0-64),"; Gain
                  Print #55, "Gamma (0-20),"; Gamma
                  Print #55, "Contrast(0-100),"; Contrast
                  Print #55, "Brightness(0-100),"; Bright
                  Print #55, "Sharpness (0-3),"; Sharp
            Close #55
End Sub
Private Sub start_Click()
    Form2.Show
    video = True
    BUFUSB_StartFrameGrab 34952     'grab frames forever
        
        Dim j As Long
        
        j = FrameCount + 1
        Do While video = True
            DoEvents
            If FrameCount >= j Then     'Don't displayimage until there's another frame, but don't crash if one is missed
            Call DisplayImage
            j = FrameCount + 1
        End If
    Loop

End Sub
Private Sub stop_Click()
    video = False
    BUFUSB_StopFrameGrab
End Sub
Private Sub BrightnessSlider_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    Bright = Value
    Call BUFUSB_SetGamma(Gamma, Contrast, Bright, Sharp)
    BrightnessSlider.Caption = "Brightness (" & Bright & ")"
End Sub
Private Sub ContrastSlider_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    Contrast = Value
    Call BUFUSB_SetGamma(Gamma, Contrast, Bright, Sharp)
    ContrastSlider.Caption = "Contrast (" & Contrast & ")"
End Sub
Private Sub ExposureSlider_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    Exposure = Value
    BUFUSB_SetExposureTime 1, Exposure / 0.05
    ExposureSlider.Caption = "Exposure Time (" & Exposure & " ms)"
End Sub
Private Sub BlackSlider_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    blacklevel = Value
    BlackSlider.Caption = "Black Cutoff (" & blacklevel & ")"
End Sub
Private Sub SharpnessSlider_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    Sharp = Value
    SharpnessSlider.Caption = "Sharpness (" & Sharp & ")"
    Call BUFUSB_SetGamma(Gamma, Contrast, Bright, Sharp)
End Sub
Private Sub GainSlider_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    Gain = Value
    BUFUSB_SetGains 1, Gain, Gain, Gain
    GainSlider.Caption = "Gain (" & Gain & ")"
End Sub
Private Sub GammaSlider_PointerValueChanged(ByVal Pointer As Long, Value As Variant)
    Gamma = Value
    Call BUFUSB_SetGamma(Gamma, Contrast, Bright, Sharp)
    GammaSlider.Caption = "Gamma (" & Gamma & ")"
End Sub

Private Sub Timer1_Timer()

If BUFUSB_InitDevice Then
        Timer1.Enabled = False
        Dim ModuleName As String
        Dim SerialName As String
        ModuleName = Space(32)
        SerialName = Space(32)
        BUFUSB_GetModuleNoSerialNo 1, ModuleName, SerialName
        ModuleNoLabel.Caption = "ModuleNo.: " & ModuleName
        BUFUSB_AddDeviceToWorkingSet 1
        BUFUSB_StartCameraEngine hWnd, 8
        
        CameraSettingsFromFile
        BUFUSB_SetCustomizedResolution 1, 1280, 1024, 0, 4
        Call BUFUSB_SetBWMode(1, 1, 0) 'flips the image if you're in BMP mode
        BUFUSB_InstallFrameHooker BMPDATA_IMAGE, AddressOf MyFrameHooker
        
        blacklevel = 128
        BlackSlider.Caption = "Black Cutoff (" & 128 & ")"
        
        Form2.Show
        video = True
        BUFUSB_StartFrameGrab 34952     'grab frames forever
        
        Sleep (100)                      'needed for some unknown reason
        
        Dim j As Long
        
        j = FrameCount + 1
        Do While video = True
            DoEvents
            If FrameCount >= j Then     'Don't displayimage until there's another frame, but don't crash if one is missed
                Call DisplayImage
                j = FrameCount + 1
            End If
        Loop
    Else
        ModuleNoLabel.Caption = "No Camera!"
End If

End Sub
