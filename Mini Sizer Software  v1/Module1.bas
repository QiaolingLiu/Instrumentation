Attribute VB_Name = "Module1"
    Public fso As New FileSystemObject
    
    Public Const e = 1.6 * 10 ^ (-19)
    Public Const Pi As Double = 3.1415926535
    Public Const A1 = 1.142
    Public Const B1 = 0.558
    Public Const B2 = -0.4995
    Public Const DMA_Length = 3.350625          'unit:cm
    Public Const DMA_Width = 2.54               'unit:cm
    Public Const DMA_Height = 0.3175            'unit:cm
    Public Const gas_viscosity = 0.0000181      'Pa s
    Public Const gas_meanfreepath = 0.0000000652  'm
    Public Const particle_density = 1.2         'g/cc
    Public Const size_max = 100
    Public Const size_min = 10   'Max/Min size range
    Public Const Volt_min = 10
    Public Const Volt_max = 5000
    Public Const Aerosol_Flow = 0.3               'lpm
    Public Const Sheath_Flow = 3                  'lpm
    Public Const Picsub_YLmax = 3                 'lpm
    Public Const Picsub_YLmin = 0                 'lpm
    Public Const Picsub_YRmax = 5000              'V
    Public Const Picsub_YRmin = 0                 'V

   'Parameters & Settings:
    Public settingflag As Boolean
    
    Public Type Steppingset
        size As Single
        Voltage As Single
        t As Integer
        tacu As Integer
    End Type
    Public Stepping() As Steppingset, SteppingNum As Integer
 
    Public cycletime As Integer, Cycle_Num As Integer   'Up/Down/Cycle Scan time
    Public size_Down As Single, size_Up As Single    'Choosen Up/Down size range
           
    Public Sample_period As Integer, Start_type$, Start_time As Date, Next_time As Date, End_time As Date, Cycle_times As Integer    'Sample_period: m
    Public Datafilepath As String
    Public Picmain_Ymax As Double, Picmain_Ymin As Double
      
    Public Xaxis_range As Integer, Sample_Num As Integer, Timersub_time_last As Date
    Public Timer_num As Long, SecNum As Long

    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
'Functions:Solve for V-->Dp

Public Function Volt_to_Dp(ByVal volt As Single) As Single

  Dim x1 As Double, x0 As Double, f As Double, f1 As Double, Cc1 As Double, Cc As Single, Q As Double, Kn As Single
  
    x1 = 0.00000005
    Q = 1 * e                   'C
    
    Do While Abs((x1 - x0)) > 0.00000000000001
        x0 = x1
        Kn = 2 * gas_meanfreepath / x0
        Cc = 1 + Kn * A1 + Kn * B1 * Exp(B2 / Kn)
        '导数
        Cc1 = -A1 * 2 * gas_meanfreepath / x0 / x0 - 2 * gas_meanfreepath / x0 / x0 * B1 * Exp(B2 / Kn) - Kn * B1 * Exp(B2 / Kn) * 2 * gas_meanfreepath / x0 / x0
        f = 3 * ((Sheath_Flow - Aerosol_Flow) / 1000 / 60) * DMA_Height * gas_viscosity * x0 * Pi / DMA_Width / (DMA_Length / 100) / Q / Cc - volt
        '导数
        f1 = 3 * ((Sheath_Flow - Aerosol_Flow) / 1000 / 60) * DMA_Height * gas_viscosity * Pi / DMA_Width / (DMA_Length / 100) / Q * (1 / Cc - x0 / Cc / Cc * Cc1)
        x1 = x0 - f / f1
    Loop
    Volt_to_Dp = Round(x1 * (10 ^ 9), 3)   'nm
    
End Function

'Functions:Solve for Dp-->V

Public Function Dp_to_Volt(ByVal Dp As Single) As Single   'Dp:nm

 Dim Cc As Single, V As Single, Q As Double, Kn As Single
      
    Dp = Dp / (10 ^ 9)
    Kn = 2 * gas_meanfreepath / Dp
    Cc = 1 + Kn * A1 + Kn * B1 * Exp(B2 / Kn)
    Q = 1 * e                   'C
    
    V = 3 * ((Sheath_Flow - Aerosol_Flow) / 1000 / 60) * DMA_Height * gas_viscosity * Dp * Pi / DMA_Width / (DMA_Length / 100) / Q / Cc   'V
    Dp_to_Volt = Round(V, 3)

End Function

Public Sub Refresh_PicSub(ByVal xrange As Integer)
    
  Dim i As Integer, j As Integer, yl_max As Single, yl_min As Single, yr_max As Single, yr_min As Single

    Frmmain.Pic_Sub.AutoRedraw = True: Frmmain.Pic_Sub.Cls: Frmmain.Pic_Sub.DrawWidth = 1
  'Y-axis
    For i = 0 To 5
        If i > 0 Then
            Frmmain.Pic_Sub.DrawStyle = 2
            Frmmain.Pic_Sub.Line (0, Frmmain.Pic_Sub.Height - (200 + 800 * i))-(Frmmain.Pic_Sub.Width, Frmmain.Pic_Sub.Height - (200 + 800 * i)), vbGrayText
        Else
            Frmmain.Pic_Sub.DrawStyle = 0
            Frmmain.Pic_Sub.Line (0, Frmmain.Pic_Sub.Height - (200 + 800 * i))-(Frmmain.Pic_Sub.Width, Frmmain.Pic_Sub.Height - (200 + 800 * i)), &H808080
        End If
    Next i
    Frmmain.Pic_Sub.DrawStyle = 0
    Frmmain.Pic_Sub.Line (60, 0)-(60, Frmmain.Pic_Sub.Height), &H808080
    Frmmain.Pic_Sub.DrawStyle = 2
  'X-axis
    If xrange > 0 Then
        For i = 1 To Int(xrange / 15) + 1
            Frmmain.Pic_Sub.Line (60 + (Frmmain.Pic_Sub.Width - 60) / (Int(xrange / 15) + 1) * i, 0)-(60 + (Frmmain.Pic_Sub.Width - 60) / (Int(xrange / 15) + 1) * i, Frmmain.Pic_Sub.Height), vbGrayText
        Next i
    End If
  
  'Sheath Flow:
    Frmmain.lbl_picsub_yl5.Caption = VBA.Format(Picsub_YLmax, "#.00")
    Frmmain.lbl_picsub_yl4.Caption = VBA.Format((Picsub_YLmax - Picsub_YLmin) * 0.8 + Picsub_YLmin, "#.00")
    Frmmain.lbl_picsub_yl3.Caption = VBA.Format((Picsub_YLmax - Picsub_YLmin) * 0.6 + Picsub_YLmin, "#.00")
    Frmmain.lbl_picsub_yl2.Caption = VBA.Format((Picsub_YLmax - Picsub_YLmin) * 0.4 + Picsub_YLmin, "#.00")
    Frmmain.lbl_picsub_yl1.Caption = VBA.Format((Picsub_YLmax - Picsub_YLmin) * 0.2 + Picsub_YLmin, "#.00")
    Frmmain.lbl_picsub_yl0.Caption = VBA.Format(Picsub_YLmin, "#0.00")
  'DMA Voltage:
    Frmmain.lbl_picsub_yr5.Caption = VBA.Format(Picsub_YRmax, "#")
    Frmmain.lbl_picsub_yr4.Caption = VBA.Format((Picsub_YRmax - Picsub_YRmin) * 0.8 + Picsub_YRmin, "#")
    Frmmain.lbl_picsub_yr3.Caption = VBA.Format((Picsub_YRmax - Picsub_YRmin) * 0.6 + Picsub_YRmin, "#")
    Frmmain.lbl_picsub_yr2.Caption = VBA.Format((Picsub_YRmax - Picsub_YRmin) * 0.4 + Picsub_YRmin, "#")
    Frmmain.lbl_picsub_yr1.Caption = VBA.Format((Picsub_YRmax - Picsub_YRmin) * 0.2 + Picsub_YRmin, "#")
    Frmmain.lbl_picsub_yr0.Caption = VBA.Format(Picsub_YRmin, "#0")
  
  'Objective Curves
    Frmmain.Pic_Sub.DrawStyle = 0
  'DMA-Voltage Setting Curve
    If settingflag = True Then
        For i = 1 To SteppingNum
            If i = 1 Then
                Frmmain.Pic_Sub.Line (60, Frmmain.Pic_Sub.Height - 200)-(60, Frmmain.Pic_Sub.Height - (200 + (Stepping(i).Voltage - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000)), vbBlue
                Frmmain.Pic_Sub.Line (60, Frmmain.Pic_Sub.Height - (200 + (Stepping(i).Voltage - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000))-(60 + Stepping(i).tacu * (Frmmain.Pic_Sub.Width - 60) / (Int(xrange / 15) + 1) / 15, Frmmain.Pic_Sub.Height - (200 + (Stepping(i).Voltage - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000)), vbBlue
            Else
                Frmmain.Pic_Sub.Line (60 + Stepping(i - 1).tacu * (Frmmain.Pic_Sub.Width - 60) / (Int(xrange / 15) + 1) / 15, Frmmain.Pic_Sub.Height - (200 + (Stepping(i - 1).Voltage - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000))-(60 + Stepping(i - 1).tacu * (Frmmain.Pic_Sub.Width - 60) / (Int(xrange / 15) + 1) / 15, Frmmain.Pic_Sub.Height - (200 + (Stepping(i).Voltage - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000)), vbBlue
                Frmmain.Pic_Sub.Line (60 + Stepping(i - 1).tacu * (Frmmain.Pic_Sub.Width - 60) / (Int(xrange / 15) + 1) / 15, Frmmain.Pic_Sub.Height - (200 + (Stepping(i).Voltage - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000))-(60 + Stepping(i).tacu * (Frmmain.Pic_Sub.Width - 60) / (Int(xrange / 15) + 1) / 15, Frmmain.Pic_Sub.Height - (200 + (Stepping(i).Voltage - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000)), vbBlue
            End If
        Next i
        Frmmain.Pic_Sub.Line (60 + Stepping(i - 1).tacu * (Frmmain.Pic_Sub.Width - 60) / (Int(xrange / 15) + 1) / 15, Frmmain.Pic_Sub.Height - (200 + (Stepping(SteppingNum).Voltage - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000))-(60 + Stepping(i - 1).tacu * (Frmmain.Pic_Sub.Width - 60) / (Int(xrange / 15) + 1) / 15, Frmmain.Pic_Sub.Height - 200), vbBlue
   End If
   
End Sub

Public Sub Refresh_PicMain(ByVal xrange As Integer)
    
  Dim i As Integer, j As Integer
  
    Frmmain.Pic_Main.AutoRedraw = True: Frmmain.Pic_Main.Cls: Frmmain.Pic_Main.DrawWidth = 1
  'Y-axis /'log
'    For i = 0 To 5
'        If i > 0 Thent
'            Frmmain.Pic_Main.DrawStyle = 2
'            Frmmain.Pic_Main.Line (0, Frmmain.Pic_Main.Height - 100 - 1000 * i)-(Frmmain.Pic_Main.Width, Frmmain.Pic_Main.Height - 100 - 1000 * i), vbGrayText
'            Frmmain.Pic_Main.DrawStyle = 0
'            For j = 2 To 9
'                Frmmain.Pic_Main.Line (60, Frmmain.Pic_Main.Height - 100 - 1000 * (i - 1) - Log(j) / Log(10) * 1000)-(150, Frmmain.Pic_Main.Height - 100 - 1000 * (i - 1) - Log(j) / Log(10) * 1000), &H808080
'            Next j
'        Else
'            Frmmain.Pic_Main.DrawStyle = 0
'            Frmmain.Pic_Main.Line (0, Frmmain.Pic_Main.Height - 100 - 1000 * i)-(Frmmain.Pic_Main.Width, Frmmain.Pic_Main.Height - 100 - 1000 * i), &H808080
'        End If
'    Next i
  'Y-axis
    For i = 0 To 5
        If i > 0 Then
            Frmmain.Pic_Main.DrawStyle = 2
            Frmmain.Pic_Main.Line (0, Frmmain.Pic_Main.Height - 100 - 1000 * i)-(Frmmain.Pic_Main.Width, Frmmain.Pic_Main.Height - 100 - 1000 * i), vbGrayText
            Frmmain.Pic_Main.DrawStyle = 0
            For j = 1 To 9
                Frmmain.Pic_Main.Line (60, Frmmain.Pic_Main.Height - 100 - 1000 * (i - 1) - j * 100)-(150, Frmmain.Pic_Main.Height - 100 - 1000 * (i - 1) - j * 100), &H808080
            Next j
        Else
            Frmmain.Pic_Main.DrawStyle = 0
            Frmmain.Pic_Main.Line (0, Frmmain.Pic_Main.Height - 100 - 1000 * i)-(Frmmain.Pic_Main.Width, Frmmain.Pic_Main.Height - 100 - 1000 * i), &H808080
        End If
    Next i
    
    Frmmain.lbl_picmain_y5.Caption = VBA.Format(Picmain_Ymax, "#.00e+0")
    Frmmain.lbl_picmain_y4.Caption = VBA.Format((Picmain_Ymax - Picmain_Ymin) * 0.8 + Picmain_Ymin, "#.00e+0")
    Frmmain.lbl_picmain_y3.Caption = VBA.Format((Picmain_Ymax - Picmain_Ymin) * 0.6 + Picmain_Ymin, "#.00e+0")
    Frmmain.lbl_picmain_y2.Caption = VBA.Format((Picmain_Ymax - Picmain_Ymin) * 0.4 + Picmain_Ymin, "#.00e+0")
    Frmmain.lbl_picmain_y1.Caption = VBA.Format((Picmain_Ymax - Picmain_Ymin) * 0.2 + Picmain_Ymin, "#.00e+0")
    Frmmain.lbl_picmain_y0.Caption = VBA.Format(Picmain_Ymin, "#.00e+0")
    
    Frmmain.Pic_Main.DrawStyle = 0
    Frmmain.Pic_Main.Line (60, 0)-(60, Frmmain.Pic_Main.Height), &H808080
    Frmmain.Pic_Main.DrawStyle = 2
  'X-axis
    If xrange > 0 Then
        For i = 1 To Int(xrange / 15) + 1
            Frmmain.Pic_Main.Line (60 + (Frmmain.Pic_Main.Width - 60) / (Int(xrange / 15) + 1) * i, 0)-(60 + (Frmmain.Pic_Main.Width - 60) / (Int(xrange / 15) + 1) * i, Frmmain.Pic_Main.Height), vbGrayText
        Next i
    End If
                
End Sub
