VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{1752FF26-D6C9-4BC8-BFE9-7D0CA26DED89}#1.0#0"; "BDaqOcx.dll"
Begin VB.Form Frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mini e-UPS"
   ClientHeight    =   10650
   ClientLeft      =   1485
   ClientTop       =   -1890
   ClientWidth     =   19080
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   19080
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame4 
      Caption         =   " Test Information "
      Height          =   2900
      Left            =   60
      TabIndex        =   4
      Top             =   400
      Width           =   9800
      Begin VB.CommandButton Cmd_shutdown 
         Caption         =   "Shut down"
         Height          =   500
         Left            =   8280
         TabIndex        =   58
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox Txt_size_range_U 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   6720
         TabIndex        =   26
         Top             =   1682
         Width           =   800
      End
      Begin VB.TextBox Txt_status 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00FF00FF&
         Height          =   400
         Left            =   5745
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   475
         Width           =   1800
      End
      Begin VB.TextBox Txt_size_range_D 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5520
         TabIndex        =   21
         Top             =   1682
         Width           =   800
      End
      Begin VB.TextBox Txt_Starttime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1665
         TabIndex        =   9
         Top             =   1682
         Width           =   2160
      End
      Begin VB.TextBox Txt_samplename 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1900
         TabIndex        =   0
         Top             =   475
         Width           =   2300
      End
      Begin VB.CommandButton Cmd_start 
         Caption         =   "Start"
         Height          =   500
         Left            =   8280
         TabIndex        =   7
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton Cmd_stop 
         Caption         =   "Stop"
         Height          =   500
         Left            =   8280
         TabIndex        =   6
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Timer Timermain 
         Left            =   3240
         Top             =   120
      End
      Begin VB.TextBox Txt_datafile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1900
         TabIndex        =   5
         Top             =   1060
         Width           =   5640
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2760
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Txt_volt_range_U 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   6720
         TabIndex        =   25
         Top             =   2302
         Width           =   800
      End
      Begin VB.TextBox Txt_volt_range_D 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5520
         TabIndex        =   19
         Top             =   2302
         Width           =   800
      End
      Begin VB.TextBox Txt_timeleft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1900
         TabIndex        =   8
         Top             =   2302
         Width           =   1200
      End
      Begin BDaqOcxLibCtl.InstantAiCtrl InstantAiCtrl1 
         Left            =   6720
         OleObjectBlob   =   "Frmmain.frx":0000
         Top             =   120
      End
      Begin BDaqOcxLibCtl.InstantAoCtrl InstantAoCtrl1 
         Left            =   6120
         OleObjectBlob   =   "Frmmain.frx":0136
         Top             =   120
      End
      Begin BDaqOcxLibCtl.InstantDoCtrl InstantDoCtrl1 
         Left            =   5280
         OleObjectBlob   =   "Frmmain.frx":01E7
         Top             =   120
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "nm"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7650
         TabIndex        =   30
         Top             =   1755
         Width           =   300
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "~"
         Height          =   255
         Left            =   6400
         TabIndex        =   28
         Top             =   2375
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "~"
         Height          =   255
         Left            =   6400
         TabIndex        =   27
         Top             =   1755
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Status :"
         Height          =   360
         Left            =   4800
         TabIndex        =   24
         Top             =   495
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "Size Range :"
         Height          =   360
         Left            =   4080
         TabIndex        =   22
         Top             =   1720
         Width           =   1860
      End
      Begin VB.Label Label38 
         Caption         =   "Time Left (s) :"
         Height          =   360
         Left            =   255
         TabIndex        =   13
         Top             =   2322
         Width           =   1500
      End
      Begin VB.Label Label14 
         Caption         =   "Start Time:"
         Height          =   360
         Left            =   255
         TabIndex        =   12
         Top             =   1720
         Width           =   1500
      End
      Begin VB.Label lbldatafile 
         Caption         =   "Data File :"
         Height          =   360
         Left            =   255
         TabIndex        =   11
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label label7 
         Caption         =   "Sample Name:"
         Height          =   360
         Left            =   255
         TabIndex        =   10
         Top             =   495
         Width           =   1545
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7650
         TabIndex        =   29
         Top             =   2375
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Voltage Range :"
         Height          =   360
         Left            =   3720
         TabIndex        =   20
         Top             =   2325
         Width           =   1860
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   56
      Top             =   10275
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13230
            MinWidth        =   13230
            Text            =   "Developed by                                                                                            Qiaoling Liu @ VCU"
            TextSave        =   "Developed by                                                                                            Qiaoling Liu @ VCU"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Version 1.0"
            TextSave        =   "Version 1.0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "9/18/2015"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   " Graph"
      Height          =   9920
      Left            =   9920
      TabIndex        =   14
      Top             =   400
      Width           =   9100
      Begin VB.PictureBox Pic_Sub 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000008&
         Height          =   4400
         Left            =   100
         ScaleHeight     =   4365
         ScaleWidth      =   8865
         TabIndex        =   17
         Top             =   5450
         Width           =   8900
         Begin VB.TextBox Txt_Sheath_Flow 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   330
            Left            =   600
            TabIndex        =   57
            Text            =   "Q_sh / lpm: "
            Top             =   500
            Width           =   1350
         End
         Begin VB.TextBox Txt_Aerosol_Flow 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   330
            Left            =   600
            TabIndex        =   32
            Text            =   "Q_a / lpm: "
            Top             =   120
            Width           =   1350
         End
         Begin VB.TextBox Txt_Scan_Volt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6840
            TabIndex        =   31
            Text            =   "Volt / V:  "
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label lbl_picsub_yr0 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   8400
            TabIndex        =   54
            Top             =   3980
            Width           =   405
         End
         Begin VB.Label lbl_picsub_yl0 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   100
            TabIndex        =   49
            Top             =   3980
            Width           =   405
         End
         Begin VB.Label lbl_picsub_yl1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0.6"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   100
            TabIndex        =   48
            Top             =   3450
            Width           =   405
         End
         Begin VB.Label lbl_picsub_yr5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "5000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   8400
            TabIndex        =   42
            Top             =   250
            Width           =   405
         End
         Begin VB.Label lbl_picsub_yr1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   8400
            TabIndex        =   41
            Top             =   3450
            Width           =   405
         End
         Begin VB.Label lbl_picsub_yr3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   8400
            TabIndex        =   40
            Top             =   1850
            Width           =   405
         End
         Begin VB.Label lbl_picsub_yr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   8400
            TabIndex        =   39
            Top             =   2650
            Width           =   405
         End
         Begin VB.Label lbl_picsub_yr4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "4000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   8400
            TabIndex        =   38
            Top             =   1050
            Width           =   405
         End
         Begin VB.Label lbl_picsub_yl5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   100
            TabIndex        =   37
            Top             =   250
            Width           =   400
         End
         Begin VB.Label lbl_picsub_yl3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1.8"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   100
            TabIndex        =   36
            Top             =   1850
            Width           =   400
         End
         Begin VB.Label lbl_picsub_yl2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1.2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   100
            TabIndex        =   35
            Top             =   2650
            Width           =   400
         End
         Begin VB.Label lbl_picsub_yl4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2.4"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   100
            TabIndex        =   34
            Top             =   1050
            Width           =   400
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Time / s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   33
            Top             =   4080
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "DMA Volts and Sheath Flow Monitoring"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   8900
         End
      End
      Begin VB.PictureBox Pic_Main 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5100
         Left            =   100
         ScaleHeight     =   5070
         ScaleWidth      =   8865
         TabIndex        =   15
         Top             =   300
         Width           =   8900
         Begin VB.TextBox Txt_Elec_data 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7560
            TabIndex        =   50
            Text            =   "Con.: "
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "dN"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   53
            Top             =   45
            Width           =   495
         End
         Begin VB.Label lbl_picmain_y0 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1.2e1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   200
            TabIndex        =   52
            Top             =   4800
            Width           =   600
         End
         Begin VB.Label lbl_picmain_y1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1.3e1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   51
            Top             =   4050
            Width           =   600
         End
         Begin VB.Label lbl_picmain_y5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1.6e1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   200
            TabIndex        =   46
            Top             =   40
            Width           =   600
         End
         Begin VB.Label lbl_picmain_y3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1.4e1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   200
            TabIndex        =   45
            Top             =   2050
            Width           =   600
         End
         Begin VB.Label lbl_picmain_y2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1.3e1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   200
            TabIndex        =   44
            Top             =   3050
            Width           =   600
         End
         Begin VB.Label lbl_picmain_y4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1.5e1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   200
            TabIndex        =   43
            Top             =   1050
            Width           =   600
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Time / s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8040
            TabIndex        =   47
            Top             =   4680
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Particle Size Distribution"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   16
            Top             =   0
            Width           =   2475
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Data "
      Height          =   7050
      Left            =   60
      TabIndex        =   1
      Top             =   3380
      Width           =   9800
      Begin VB.ListBox List_Sample 
         Height          =   5475
         ItemData        =   "Frmmain.frx":0245
         Left            =   150
         List            =   "Frmmain.frx":0247
         TabIndex        =   65
         Top             =   1465
         Width           =   3705
      End
      Begin VB.TextBox Txt_temp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1440
         TabIndex        =   62
         Top             =   400
         Width           =   750
      End
      Begin VB.TextBox Txt_HR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3480
         TabIndex        =   61
         Top             =   400
         Width           =   750
      End
      Begin VB.TextBox Txt_charger_p 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   6600
         TabIndex        =   59
         Top             =   400
         Width           =   800
      End
      Begin MSFlexGridLib.MSFlexGrid MSFL_results 
         Height          =   1905
         Left            =   4005
         TabIndex        =   66
         Top             =   5035
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   3360
         _Version        =   393216
         Rows            =   5
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSFL_data 
         Height          =   3555
         Left            =   4005
         TabIndex        =   67
         Top             =   1465
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   6271
         _Version        =   393216
         Rows            =   5
         ScrollTrack     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   9800
         Y1              =   950
         Y2              =   950
      End
      Begin VB.Label Label17 
         Caption         =   "Temp (C) :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   64
         Top             =   460
         Width           =   1140
      End
      Begin VB.Label Label16 
         Caption         =   "HR (%) :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   63
         Top             =   460
         Width           =   1020
      End
      Begin VB.Label Label12 
         Caption         =   "Charger +H.V. (kV) :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         TabIndex        =   60
         Top             =   460
         Width           =   2340
      End
      Begin VB.Label Label40 
         Caption         =   "Data and Results :"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   1100
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "Sample List :"
         Height          =   375
         Left            =   200
         TabIndex        =   3
         Top             =   1100
         Width           =   2055
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   25
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   16
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   17
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   18
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   19
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   20
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   21
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   22
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10080
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   22
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":0249
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":129B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":22ED
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":333F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":4391
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":53E3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":6435
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":7487
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":84D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":952B
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":A90D
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":BCEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":D0D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":E123
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":F175
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":101C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":11219
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":1226B
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":132BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":1430F
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":15361
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmmain.frx":163B3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File(F)"
      Begin VB.Menu Mnunew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Mnusaveas 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu Mnupics 
         Caption         =   "Export to &Pics"
      End
      Begin VB.Menu Mnuprint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MnuRun 
      Caption         =   "Run(R)"
      Begin VB.Menu Mnustart 
         Caption         =   "&Start"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Mnustop 
         Caption         =   "&Stop"
      End
   End
   Begin VB.Menu MnuSettings 
      Caption         =   "Properties(P)"
      Begin VB.Menu MnuPS 
         Caption         =   "Properties and Settings"
      End
   End
   Begin VB.Menu Mnusample 
      Caption         =   "Sample(S)"
      Begin VB.Menu Mnulast 
         Caption         =   "Last Sample"
      End
      Begin VB.Menu Mnunext 
         Caption         =   "Next Sample"
      End
      Begin VB.Menu Mnufirst 
         Caption         =   "First Sample"
      End
      Begin VB.Menu Mnuending 
         Caption         =   "Ending Sample"
      End
   End
   Begin VB.Menu Mnuview 
      Caption         =   "Data(D)"
      Begin VB.Menu Mnuweight 
         Caption         =   "Weight"
         Begin VB.Menu Mnuweight1 
            Caption         =   "Number (N-#/cm3)"
            Checked         =   -1  'True
         End
         Begin VB.Menu Mnuweight2 
            Caption         =   "Surface (S um2/cm3)"
         End
         Begin VB.Menu Mnuweight3 
            Caption         =   "Volume (V um3/cm3)"
         End
         Begin VB.Menu Mnuweight4 
            Caption         =   "Mass (m ug/cm3)"
         End
         Begin VB.Menu Mnuunit6 
            Caption         =   "Counts(#)"
         End
      End
      Begin VB.Menu Mnuunit 
         Caption         =   "Unit"
         Begin VB.Menu Mnuunit2 
            Caption         =   "dW (#/cm3)"
            Checked         =   -1  'True
         End
         Begin VB.Menu Mnuunit1 
            Caption         =   "dW/dlogDp (#/cm3)"
         End
         Begin VB.Menu Mnuunit3 
            Caption         =   "dW/W %"
         End
         Begin VB.Menu Mnuunit5 
            Caption         =   "Accumulative Weight"
         End
         Begin VB.Menu Mnuunit4 
            Caption         =   "Acummulative %"
         End
      End
   End
   Begin VB.Menu Mnuplot 
      Caption         =   "Plot(P)"
      Begin VB.Menu Mnutype 
         Caption         =   "Type"
         Begin VB.Menu Mnuspot 
            Caption         =   "Spot"
         End
         Begin VB.Menu Mnuline 
            Caption         =   "Line"
         End
         Begin VB.Menu Mnubar 
            Caption         =   "Bar"
            Checked         =   -1  'True
         End
         Begin VB.Menu Mnuarea 
            Caption         =   "Area"
         End
      End
      Begin VB.Menu Mnugridline 
         Caption         =   "Gridline"
         Begin VB.Menu Mnumajorv 
            Caption         =   "Major Vertical"
         End
         Begin VB.Menu Mnuminorv 
            Caption         =   "Minor Vertical"
         End
         Begin VB.Menu Mnumajorh 
            Caption         =   "Major Horizontal"
         End
         Begin VB.Menu Mnuminorh 
            Caption         =   "Minor Horizontal"
         End
      End
      Begin VB.Menu Mnucolor 
         Caption         =   "Color"
      End
      Begin VB.Menu MnuYaxis 
         Caption         =   "Y-axis"
         Begin VB.Menu MnuYcommon 
            Caption         =   "Common"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuYlog 
            Caption         =   "Log"
         End
      End
      Begin VB.Menu MnuXaxis 
         Caption         =   "X-axis"
         Begin VB.Menu MnuXcommon 
            Caption         =   "Common"
         End
         Begin VB.Menu MnuXlog 
            Caption         =   "Log"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu MnuWindows 
      Caption         =   "Windows(W)"
      Begin VB.Menu MnuSettingWindow 
         Caption         =   "Setting"
      End
      Begin VB.Menu MnuMainWindow 
         Caption         =   "Main Window"
      End
      Begin VB.Menu MnuOffline 
         Caption         =   "Data Reduction (offline)"
      End
   End
   Begin VB.Menu Mnuhelp1 
      Caption         =   "Help(H)"
      Begin VB.Menu Mnuhelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Status$
    Private Timermain_time_last As Date

    Private AveData_Elec As Single, SumData_Elec As Double, AveData_Elec_array() As Double
    Private Avedata_Aerosol_Flow As Single, Sumdata_Aerosol_Flow As Double, last_Aerosol_Flow As Single
    Private Avedata_Sheath_Flow As Single, Sumdata_Sheath_Flow As Double, last_Sheath_Flow As Single
    Private Avedata_DMA_Volt As Single, Sumdata_DMA_Volt As Double, last_DMA_Volt As Single
    Private Avedata_Charger_P_Volt As Single, Sumdata_Charger_P_Volt As Double
    Private Avedata_HR As Single, Sumdata_HR As Double
    Private Avedata_Temp As Single, Sumdata_Temp As Double
    
    Private Step As Integer
        
    Private Data_max As Double, Data_min As Double
    
    Private end_flag As Boolean
    
    Private WfileA As String, WfileB As String
'DAQ:

    Private Const CHANNEL_COUNT_MAX As Integer = 16
    Private dataScaled_AI(CHANNEL_COUNT_MAX - 1) As Double
    Private Const chanCountSet As Integer = 8
    Private Const StartPort As Integer = 0
    Private Const PortCountShow As Integer = 4
    Private chanCountMax As Integer
    Private portDatas As Byte
    Private Const channelStart As Integer = 0
    Private Const channelCount As Integer = 1
    Private dataScaled_AO(0 To 1) As Double
        

 
Private Sub Form_Load()
 
Dim i As Integer
 
  On Error Resume Next
  
    Timermain.Enabled = True: Timermain.Interval = 50
    Cmd_start.Enabled = False: Cmd_stop.Enabled = False: Mnustart.Enabled = False: Mnustop.Enabled = False
    
    MSFL_data.Cols = 7: MSFL_data.Rows = 20
    MSFL_data.WordWrap = True:  MSFL_data.HighLight = flexHighlightNever
    MSFL_data.ColWidth(0) = 500
    For i = 1 To MSFL_data.Cols - 1
        MSFL_data.ColWidth(i) = 1000
    Next i
    MSFL_data.RowHeight(0) = 550
    For i = 1 To MSFL_data.Rows - 1
        MSFL_data.RowHeight(i) = 360
    Next i
    MSFL_data.TextMatrix(0, 0) = "NO.": MSFL_data.TextMatrix(0, 1) = "Size(mid) nm"
    MSFL_data.TextMatrix(0, 2) = "Con #/cm3": MSFL_data.TextMatrix(0, 3) = "Surface um2/cm3"
    MSFL_data.TextMatrix(0, 4) = "Volume um3/cm3": MSFL_data.TextMatrix(0, 5) = "Mass ug/cm3"
    MSFL_data.TextMatrix(0, 6) = "Volt V"
    
    MSFL_results.Cols = 5: MSFL_results.Rows = 7
    MSFL_results.WordWrap = True:  MSFL_results.HighLight = flexHighlightNever
    MSFL_results.ColWidth(0) = 1335
    For i = 1 To MSFL_results.Cols - 1
        MSFL_results.ColWidth(i) = 1000
    Next i
    For i = 0 To MSFL_results.Rows - 1
        MSFL_results.RowHeight(i) = 360
    Next i
    MSFL_results.TextMatrix(0, 1) = "Number": MSFL_results.TextMatrix(0, 2) = "Surface": MSFL_results.TextMatrix(0, 3) = "Volume"
    MSFL_results.TextMatrix(0, 4) = "Mass"
    MSFL_results.TextMatrix(1, 0) = "Media": MSFL_results.TextMatrix(2, 0) = "Mean": MSFL_results.TextMatrix(3, 0) = "Mode"
    MSFL_results.TextMatrix(4, 0) = "Geo.Mean": MSFL_results.TextMatrix(5, 0) = "Geo.Std. Dev.": MSFL_results.TextMatrix(6, 0) = "Total Con"
    
  'Graphs:
    Xaxis_range = 120: Picmain_Ymax = 0: Picmain_Ymin = 0

    Refresh_PicSub (Xaxis_range)
    Refresh_PicMain (Xaxis_range)
    
  'Creat Data Folder
    If fso.FolderExists(App.Path & "\Mini-Sizer Data") = False Then
        fso.CreateFolder (App.Path & "\Mini-Sizer Data")
    End If
    
    
  'DAQ Card Load:
    If Not InstantDoCtrl1.Initialized Then
        MsgBox "Please select a device with DAQNavi wizard!"
        End
    End If
    If Not InstantAiCtrl1.Initialized Then
        MsgBox "Please select a device with DAQNavi wizard!"
        End
    End If
    If Not InstantAoCtrl1.Initialized Then
        MsgBox "Please select a device with DAQNavi wizard!"
        End
    End If

    Dim devNum As Long
    Dim devDesc As String
    Dim devMode As AccessMode
    Dim modIndex As Long
    InstantDoCtrl1.getSelectedDevice devNum, devDesc, devMode, modIndex
    chanCountMax = InstantAiCtrl1.Features.ChannelCountMax
    InstantAiCtrl1.getSelectedDevice devNum, devDesc, devMode, modIndex
    InstantAoCtrl1.getSelectedDevice devNum, devDesc, devMode, modIndex
    
    InitializePortState
  
  'Pumps On
    portDatas = 3
    err = InstantDoCtrl1.WritePort(0 + StartPort, portDatas)
    If err <> Success Then
        Call HandleError(err)
        Timermain.Enabled = False
    End If

  'Initialization all parameters:
    settingflag = False
    
    Txt_samplename.Text = ""
    Status = "Standby":  Txt_status.Text = Status
    StatusBar1.Panels(3).Text = "Please set a new sample sequence"
    
    Timer_num = 0: SecNum = 0
    
    AveData_Elec = 0: SumData_Elec = 0
    Avedata_Aerosol_Flow = 0: Sumdata_Aerosol_Flow = 0: last_Aerosol_Flow = 0
    Avedata_Sheath_Flow = 0: Sumdata_Sheath_Flow = 0: last_Sheath_Flow = 0
    Avedata_DMA_Volt = 0: Sumdata_DMA_Volt = 0: last_DMA_Volt = 0
    Avedata_Charger_P_Volt = 0: Sumdata_Charger_P_Volt = 0
    Avedata_HR = 0: Sumdata_HR = 0
    Avedata_Temp = 0: Sumdata_Temp = 0
    
    SteppingNum = 0: Step = 0: Cycle_Num = 0: Sample_Num = 0: Cycle_times = 0
    
    Data_max = 0: Data_min = 0
    end_flag = False
   'Creat Data Folder
    If fso.FolderExists(App.Path & "\Mini-Sizer Data\Raw data") = False Then
        fso.CreateFolder (App.Path & "\Mini-Sizer Data\Raw data")
    End If
    WfileA = App.Path & "\Mini-Sizer Data\Raw data\A.TXT"
    WfileB = App.Path & "\Mini-Sizer Data\Raw data\B.TXT"
    fso.CreateTextFile (WfileA)
    fso.CreateTextFile (WfileB)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Dim feedback As Integer
    feedback = MsgBox("  Are you sure to Exit ?", 67, " SMPS_Software WUSTL")
  DoEvents
  On Error Resume Next
    If feedback = vbYes Then
        If end_flag = False Then
            Open Datafilepath For Append As #1
            Print #1, "End"
            Close #1
            end_flag = True
        End If
 
      'DMA Voltage Set to 0
        dataScaled_AO(0) = 0: dataScaled_AO(1) = 0
        err = InstantAoCtrl1.WriteChannels(channelStart, channelCount, dataScaled_AO)
        If err <> Success Then Call HandleError(err)
      'Turn off pumps
        portDatas = 0
        err = InstantDoCtrl1.WritePort(0 + StartPort, portDatas)
        If err <> Success Then Call HandleError(err)
        
        Timermain.Enabled = False
        Unload Frmsettings
        Unload Frmmain
        End
    Else
        Cancel = 1
    End If
    
End Sub

Private Sub Cmd_start_Click()

  Dim Timermain_time_now As Date
     
    Timermain_time_now = VBA.Now
    
    If Start_type = "Immediately" Then
        Start_time = Now + 2 / 24 / 60 / 60
        Status = "Next Sample"
        Txt_Starttime.Text = VBA.Format(Start_time, "mm-dd-yyyy  hh:mm:ss")
    Else
        If Timermain_time_now > Start_time Then
            MsgBox "Please Reset Starting time!", 48, "Mini-Sizer Software"
            Status = "Standby": Txt_Starttime.Text = ""
        End If
    End If
    
    Txt_Starttime.Text = VBA.Format(Start_time, "mm-dd-yyyy  hh:mm:ss")
    Txt_status.Text = Status: StatusBar1.Panels(3).Text = Status: Txt_status.BackColor = &HFFFFF
  
  'Refresh window:
    Txt_samplename.Text = ""
    Txt_timeleft.Text = cycletime
    
    Cmd_stop.Enabled = True: Cmd_start.Enabled = False: Cmd_shutdown.Enabled = False
    Mnustop.Enabled = True: Cmd_start.Enabled = False: MnuOpen.Enabled = False
    
    MSFL_data.Clear
    MSFL_data.TextMatrix(0, 0) = "NO.": MSFL_data.TextMatrix(0, 1) = "Size(mid) nm"
    MSFL_data.TextMatrix(0, 2) = "Con #/cm3": MSFL_data.TextMatrix(0, 3) = "Surface um2/cm3"
    MSFL_data.TextMatrix(0, 4) = "Volume um3/cm3": MSFL_data.TextMatrix(0, 5) = "Mass ug/cm3"
    MSFL_data.TextMatrix(0, 6) = "Volt V"
    
    MSFL_results.Clear
    MSFL_results.TextMatrix(0, 1) = "Number": MSFL_results.TextMatrix(0, 2) = "Surface": MSFL_results.TextMatrix(0, 3) = "Volume"
    MSFL_results.TextMatrix(0, 4) = "Mass"
    MSFL_results.TextMatrix(1, 0) = "Media": MSFL_results.TextMatrix(2, 0) = "Mean": MSFL_results.TextMatrix(3, 0) = "Mode"
    MSFL_results.TextMatrix(4, 0) = "Geo.Mean": MSFL_results.TextMatrix(5, 0) = "Geo.Std. Dev.": MSFL_results.TextMatrix(6, 0) = "Total Con"
        
    Refresh_PicSub (Xaxis_range)
    Refresh_PicMain (Xaxis_range)
    Txt_Elec_data.Text = "Con. :": Txt_aerosol_flow.Text = "Q_a / lpm :": Txt_sheath_flow.Text = "Q_sh / lpm :": Txt_Scan_Volt.Text = "Volt / V :"
    
  'Initilization parameters
    
    Timer_num = 0: SecNum = 0
    Timermain_time_last = VBA.Time
    
    AveData_Elec = 0: SumData_Elec = 0
    Avedata_Aerosol_Flow = 0: Sumdata_Aerosol_Flow = 0: last_Aerosol_Flow = 0
    Avedata_Sheath_Flow = 0: Sumdata_Sheath_Flow = 0: last_Sheath_Flow = 0
    Avedata_DMA_Volt = 0: Sumdata_DMA_Volt = 0: last_DMA_Volt = 0
    Avedata_Charger_P_Volt = 0: Sumdata_Charger_P_Volt = 0
    Avedata_HR = 0: Sumdata_HR = 0
    Avedata_Temp = 0: Sumdata_Temp = 0
    
    Data_max = 0: Data_min = 0
    end_flag = False
    
    Step = 0
    Picmain_Ymax = 0: Picmain_Ymin = 0
        
End Sub

Private Sub mnuexit_Click()
    Unload Frmmain
End Sub

Private Sub Mnunew_Click()
  DoEvents
    settingflag = False
    Txt_samplename.Text = ""
    '判断 如果timer正在运行
    Load Frmsettings
    Frmsettings.Show
End Sub


Private Sub Mnustart_Click()
    Cmd_start_Click
End Sub

Private Sub cmd_stop_Click()
  
  Dim Timermain_time_now As Date
     
    Timermain_time_now = VBA.Now
  DoEvents
    Cmd_start.Enabled = True: Cmd_stop.Enabled = False: Cmd_shutdown.Enabled = True
    Mnustart.Enabled = True:  Mnustop.Enabled = False: Mnunew.Enabled = True
    Txt_samplename.ForeColor = vbGrayText
    List_Sample.Enabled = True
    If end_flag = False Then
        Open Datafilepath For Append As #1
        Print #1, "End"
        Close #1
        end_flag = True
    End If
    
    If Cycle_Num = Cycle_times Then Cycle_Num = 0
    
    Status = "Standby": Txt_status.Text = Status: StatusBar1.Panels(3).Text = Status
    MnuPS.Enabled = True
    
  'DMA Voltage Set to 0
    dataScaled_AO(0) = 0: dataScaled_AO(1) = 0
    err = InstantAoCtrl1.WriteChannels(channelStart, channelCount, dataScaled_AO)
    If err <> Success Then Call HandleError(err)
    
End Sub

Private Sub Cmd_shutdown_Click()

  DoEvents
  On Error Resume Next
    If Cmd_shutdown.Caption = "Shut down" Then
        Cmd_shutdown.Caption = "Turn on"
        Cmd_start.Enabled = False: Cmd_stop.Enabled = False
        Mnunew.Enabled = False: Mnustart.Enabled = False:  Mnustop.Enabled = False
        If Status = "Stepping" And end_flag = False Then
            Open Datafilepath For Append As #1
            Print #1, "End"
            Close #1
            end_flag = True
        End If
        Status = "Shut Down": Txt_status.Text = Status: StatusBar1.Panels(3).Text = Status
        
      'DMA Voltage Set to 0
        dataScaled_AO(0) = 0: dataScaled_AO(1) = 0
        err = InstantAoCtrl1.WriteChannels(channelStart, channelCount, dataScaled_AO)
        If err <> Success Then Call HandleError(err)
      'Turn off pumps
        portDatas = 0
        err = InstantDoCtrl1.WritePort(0 + StartPort, portDatas)
        If err <> Success Then Call HandleError(err)
        
        
        Txt_samplename.Text = "": Txt_Starttime.Text = "": Txt_timeleft.Text = "": Txt_size_range_D.Text = "": Txt_size_range_U.Text = ""
        Txt_volt_range_D.Text = "": Txt_volt_range_U.Text = ""
        Txt_aerosol_flow.Text = "": Txt_sheath_flow.Text = "": Txt_temp.Text = "": Txt_HR.Text = "": Txt_charger_p.Text = ""
        Txt_Elec_data.Text = "Con.": Txt_Scan_Volt.Text = "Volt / V:"
        
        MSFL_data.Clear
        MSFL_data.TextMatrix(0, 0) = "NO.": MSFL_data.TextMatrix(0, 1) = "Size(mid) nm"
        MSFL_data.TextMatrix(0, 2) = "Con #/cm3": MSFL_data.TextMatrix(0, 3) = "Surface um2/cm3"
        MSFL_data.TextMatrix(0, 4) = "Volume um3/cm3": MSFL_data.TextMatrix(0, 5) = "Mass ug/cm3"
        MSFL_data.TextMatrix(0, 6) = "Volt V"
        
        MSFL_results.Clear
        MSFL_results.TextMatrix(0, 1) = "Number": MSFL_results.TextMatrix(0, 2) = "Surface": MSFL_results.TextMatrix(0, 3) = "Volume"
        MSFL_results.TextMatrix(0, 4) = "Mass"
        MSFL_results.TextMatrix(1, 0) = "Media": MSFL_results.TextMatrix(2, 0) = "Mean": MSFL_results.TextMatrix(3, 0) = "Mode"
        MSFL_results.TextMatrix(4, 0) = "Geo.Mean": MSFL_results.TextMatrix(5, 0) = "Geo.Std. Dev.": MSFL_results.TextMatrix(6, 0) = "Total Con"
        
        Refresh_PicSub (Xaxis_range)
        Refresh_PicMain (Xaxis_range)
    Else
        Cmd_shutdown.Caption = "Shut down"
        Cmd_start.Enabled = True: Cmd_stop.Enabled = True
        Mnunew.Enabled = True:  Mnustart.Enabled = True: Mnustop.Enabled = True
        Status = "Standby": Txt_status.Text = Status: StatusBar1.Panels(3).Text = Status
        
        InitializePortState
      'Turn on pumps
        portDatas = 3
        err = InstantDoCtrl1.WritePort(0 + StartPort, portDatas)
        If err <> Success Then Call HandleError(err)

    End If
    
End Sub

Private Sub Mnustop_Click()
    cmd_stop_Click
End Sub

Private Sub Mnups_Click()
    Load Frmsettings
    Frmsettings.Show
End Sub

'Main Program for Mini-Sizer System Control:
'-DMA Voltage Control: Stepping
'-Data Reading and Recording

'Status: 1 - "Standby" (for setting new sample sequence)
'        2 - "Next Sample" ("Waiting for Next Sample")
'        3 - "Stepping"
'        4 - "Shut Down"
'

  'DO port: 0-Aerosol pump,1-Sheath pump
  'AO port: 0-DMA
  'AI port: 0-Aerosol flow,1-Sheath flow,2-DMA HV,3-Charger PHV,4-Charger NHV,5-HR,6-Electrometer,7-Temp
  
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Sub Timermain_Timer()

  Dim DMA_Voltage_Output As Double, Data(0 To 9) As Double, i As Integer, Timermain_time_now As Date
     
    If Status <> "Shut Down" Then

        Timermain_time_now = VBA.Now
        
        err = InstantAiCtrl1.ReadChannels(0, chanCountSet, Null, dataScaled_AI)
        If err <> Success Then Call HandleError(err)
        
        If Timermain_time_now = Timermain_time_last Then
          'Data processing:
            Timer_num = Timer_num + 1
            Sumdata_Aerosol_Flow = SumData_Elec + dataScaled_AI(0)
            Sumdata_Sheath_Flow = Sumdata_Sheath_Flow + dataScaled_AI(1)
            Sumdata_DMA_Volt = Sumdata_DMA_Volt + dataScaled_AI(2)
            Sumdata_Charger_P_Volt = Sumdata_Charger_P_Volt + dataScaled_AI(3)
            Sumdata_HR = Sumdata_HR + dataScaled_AI(5)
            SumData_Elec = SumData_Elec + dataScaled_AI(6)
            Sumdata_Temp = Sumdata_Temp + dataScaled_AI(7)
        Else
            If Timer_num <> 0 Then
              'Averaging Data:
'+++++++++++++++++++++++++++++++++++++++++++ need calibration
                Avedata_Sheath_Flow = Round(Sumdata_Sheath_Flow / Timer_num, 2)
                Avedata_Aerosol_Flow = Round(Sumdata_Aerosol_Flow / Timer_num, 2)
                Avedata_Charger_P_Volt = Round(Sumdata_Charger_P_Volt / Timer_num * 4348.4 - 4670.6, 2)
                Avedata_DMA_Volt = Round(Sumdata_DMA_Volt / Timer_num * 4281.6 - 4570.2, 2)
                AveData_Elec = Round((SumData_Elec / Timer_num + 0.0098) * 2010.2, 2)
                Avedata_HR = Round(Sumdata_HR / Timer_num, 2)
                Avedata_Temp = Round(Sumdata_Temp / Timer_num, 2)
    
                Txt_aerosol_flow.Text = "Q_a / lpm:" & VBA.Format(Avedata_Aerosol_Flow, "#0.00")
                Txt_sheath_flow.Text = "Q_sh / lpm:" & VBA.Format(Avedata_Sheath_Flow, "#0.00")
                Txt_temp.Text = VBA.Format(Avedata_Temp, "#.00")
                Txt_HR.Text = VBA.Format(Avedata_HR, "#0.00")
                Txt_charger_p.Text = VBA.Format(Avedata_Charger_P_Volt, "#0.00")
                               
                Timer_num = 1
                Sumdata_Aerosol_Flow = dataScaled_AI(0)
                Sumdata_Sheath_Flow = dataScaled_AI(1)
                Sumdata_Charger_P_Volt = dataScaled_AI(3)
                Sumdata_HR = dataScaled_AI(5)
                Sumdata_Temp = dataScaled_AI(7)
                SumData_Elec = dataScaled_AI(6)
                Sumdata_DMA_Volt = dataScaled_AI(2)
            End If
          
          'Status & Step & SecNum Determining
          'First Second:
            If Timermain_time_now = Start_time Then
                Status = "Stepping"
                Step = 1
                SecNum = 0
                Cycle_Num = Cycle_Num + 1
                Sample_Num = Sample_Num + 1
                
                Txt_samplename.Text = "Sample " & Sample_Num
                List_Sample.AddItem Txt_samplename.Text
                List_Sample.ListIndex = List_Sample.ListCount - 1
                List_Sample.Enabled = False
                        
                MSFL_data.Clear
                MSFL_data.TextMatrix(0, 0) = "NO.": MSFL_data.TextMatrix(0, 1) = "Size(mid) nm"
                MSFL_data.TextMatrix(0, 2) = "Con #/cm3": MSFL_data.TextMatrix(0, 3) = "Surface um2/cm3"
                MSFL_data.TextMatrix(0, 4) = "Volume um3/cm3": MSFL_data.TextMatrix(0, 5) = "Mass ug/cm3"
                MSFL_data.TextMatrix(0, 6) = "Volt V"
                MSFL_results.Clear
                MSFL_results.TextMatrix(0, 1) = "Number": MSFL_results.TextMatrix(0, 2) = "Surface": MSFL_results.TextMatrix(0, 3) = "Volume"
                MSFL_results.TextMatrix(0, 4) = "Mass"
                MSFL_results.TextMatrix(1, 0) = "Media": MSFL_results.TextMatrix(2, 0) = "Mean": MSFL_results.TextMatrix(3, 0) = "Mode"
                MSFL_results.TextMatrix(4, 0) = "Geo.Mean": MSFL_results.TextMatrix(5, 0) = "Geo.Std. Dev.": MSFL_results.TextMatrix(6, 0) = "Total Con"
                
                ReDim AveData_Elec_array(1 To cycletime) As Double
                Refresh_PicSub (Xaxis_range)
                Refresh_PicMain (Xaxis_range)
                
                MnuPS.Enabled = False
            End If
            
          'Check/Control Status:
            If Status = "Stepping" And Timermain_time_now >= Start_time And SecNum >= 0 And SecNum <= cycletime And settingflag = True Then
                SecNum = SecNum + 1
                If SecNum < cycletime Then
                    If SecNum = Stepping(Step).tacu + 1 And Step < SteppingNum Then Step = Step + 1
                    Txt_timeleft.Text = cycletime - SecNum
                  'DMA Voltage:
                    DMA_Voltage_Output = Stepping(Step).Voltage
                    If DMA_Voltage_Output < 0 Then DMA_Voltage_Output = 0
                    dataScaled_AO(0) = DMA_Voltage_Output / Volt_max * 5: dataScaled_AO(1) = 0
                End If
              'Display and recording last sec data
                Txt_Elec_data.Text = "Con.: " & VBA.Format(AveData_Elec, "#.00e+00")
                Txt_Scan_Volt.Text = "Volt / V:" & VBA.Format(Avedata_DMA_Volt, "#0.00")
                If SecNum = (MSFL_data.Rows - 2) Then MSFL_data.AddItem (MSFL_data.Rows): MSFL_data.RowHeight(MSFL_data.Rows - 1) = 360
                MSFL_data.TextMatrix(SecNum, 0) = SecNum
                MSFL_data.TextMatrix(SecNum, 1) = Stepping(Step).size
                MSFL_data.TextMatrix(SecNum, 2) = AveData_Elec
                MSFL_data.TextMatrix(SecNum, 6) = VBA.Format(Avedata_DMA_Volt, "#0.00")
              'Record into File:
                If Datafilepath <> "" Then
                    Open Datafilepath For Append As #1
                    If SecNum = 1 Then
                        Print #1, Tab(1); "* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
                        Print #1, "Test Name: " & Txt_samplename.Text
                        Print #1, Tab(1); "Time"; Tab(13); "Size (nm)"; Tab(23); "DMA_Voltage (V)"; Tab(43); "Elec_Data (#/cm3)"; Tab(63); "Q_Sheath (lpm)"; _
                                  Tab(83); "Q_Aerosol (lpm)"; Tab(103); "Charger + (V)"; Tab(123); "Temp (C)"; Tab(143); "HR (%)"
                        end_flag = False
                    End If
                    Print #1, Tab(1); VBA.DateAdd("s", -1, VBA.Time); Tab(13); MSFL_data.TextMatrix(SecNum, 1); Tab(23); MSFL_data.TextMatrix(SecNum, 6); Tab(43); AveData_Elec; _
                              Tab(63); Avedata_Sheath_Flow; Tab(83); Avedata_Aerosol_Flow; ; Tab(103); Txt_charger_p.Text; _
                              Tab(123); Txt_temp.Text; Tab(143); Txt_HR.Text
                    If SecNum = cycletime + 1 And end_flag = False Then
                        Print #1, "End"
                        end_flag = True
                    End If
                    Close #1
                End If
            
'Wireless Data transmission:
                If (Sample_Num Mod 2 = 1) And WfileA <> "" Then
                    If SecNum = 1 Then
                        Open WfileA For Output As #1
                        Print #1, Tab(1); "* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
                        Print #1, "Test Name: " & Txt_samplename.Text
                        Print #1, Tab(1); "Time"; Tab(13); "Size (nm)"; Tab(23); "DMA_Voltage (V)"; Tab(43); "Elec_Data (#/cm3)"; Tab(63); "Q_Sheath (lpm)"; _
                                  Tab(83); "Q_Aerosol (lpm)"; Tab(103); "Charger + (V)"; Tab(123); "Temp (C)"; Tab(143); "HR (%)"
                        Print #1, Tab(1); VBA.DateAdd("s", -1, VBA.Time); Tab(13); MSFL_data.TextMatrix(SecNum, 1); Tab(23); MSFL_data.TextMatrix(SecNum, 6); Tab(43); AveData_Elec; _
                                  Tab(63); Avedata_Sheath_Flow; Tab(83); Avedata_Aerosol_Flow; ; Tab(103); Txt_charger_p.Text; _
                                  Tab(123); Txt_temp.Text; Tab(143); Txt_HR.Text
                    Else
                        Open WfileA For Append As #1
                        Print #1, Tab(1); VBA.DateAdd("s", -1, VBA.Time); Tab(13); MSFL_data.TextMatrix(SecNum, 1); Tab(23); MSFL_data.TextMatrix(SecNum, 6); Tab(43); AveData_Elec; _
                                  Tab(63); Avedata_Sheath_Flow; Tab(83); Avedata_Aerosol_Flow; ; Tab(103); Txt_charger_p.Text; _
                                  Tab(123); Txt_temp.Text; Tab(143); Txt_HR.Text
                        If SecNum = cycletime + 1 And end_flag = False Then
                            Print #1, "End"
                            end_flag = True
                        End If
                    End If
                    Close #1
                End If
              
                If (Sample_Num Mod 2 = 0) And WfileB <> "" Then
                    If SecNum = 1 Then
                        Open WfileB For Output As #1
                        Print #1, Tab(1); "* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
                        Print #1, "Test Name: " & Txt_samplename.Text
                        Print #1, Tab(1); "Time"; Tab(13); "Size (nm)"; Tab(23); "DMA_Voltage (V)"; Tab(43); "Elec_Data (#/cm3)"; Tab(63); "Q_Sheath (lpm)"; _
                                  Tab(83); "Q_Aerosol (lpm)"; Tab(103); "Charger + (V)"; Tab(123); "Temp (C)"; Tab(143); "HR (%)"
                        Print #1, Tab(1); VBA.DateAdd("s", -1, VBA.Time); Tab(13); MSFL_data.TextMatrix(SecNum, 1); Tab(23); MSFL_data.TextMatrix(SecNum, 6); Tab(43); AveData_Elec; _
                                  Tab(63); Avedata_Sheath_Flow; Tab(83); Avedata_Aerosol_Flow; ; Tab(103); Txt_charger_p.Text; _
                                  Tab(123); Txt_temp.Text; Tab(143); Txt_HR.Text
                    Else
                        Open WfileB For Append As #1
                        Print #1, Tab(1); VBA.DateAdd("s", -1, VBA.Time); Tab(13); MSFL_data.TextMatrix(SecNum, 1); Tab(23); MSFL_data.TextMatrix(SecNum, 6); Tab(43); AveData_Elec; _
                                  Tab(63); Avedata_Sheath_Flow; Tab(83); Avedata_Aerosol_Flow; ; Tab(103); Txt_charger_p.Text; _
                                  Tab(123); Txt_temp.Text; Tab(143); Txt_HR.Text
                        If SecNum = cycletime + 1 And end_flag = False Then
                            Print #1, "End"
                            end_flag = True
                        End If
                    End If
                    Close #1
                End If
              
              
              'Picture:
                If SecNum = 2 Then
                    Data_max = AveData_Elec: Data_min = 0
                    If Data_max > Picmain_Ymax Then Picmain_Ymax = Data_max * 1.2
                Else
                    If AveData_Elec > Data_max Then Data_max = AveData_Elec
                    If AveData_Elec < Data_min Then Data_min = AveData_Elec
                    If Data_max > Picmain_Ymax Then Picmain_Ymax = Data_max * 1.2
                    If Data_min < Picmain_Ymin Then Picmain_Ymin = Data_min * 0.9
                End If
                If SecNum > 1 Then AveData_Elec_array(SecNum - 1) = AveData_Elec
                Refresh_PicMain (Xaxis_range)
                Pic_Main.AutoRedraw = True: Pic_Main.DrawStyle = vbSolid
                Pic_Main.DrawWidth = 3
                For i = 1 To SecNum - 1
                  Dim X As Double, Y As Double
                    X = 60 + i * (Pic_Main.Width - 60) / (Int(Xaxis_range / 15) + 1) / 15
                    Y = (AveData_Elec_array(i) - Picmain_Ymin) / (Picmain_Ymax - Picmain_Ymin) * 5000
                    'y = (Log(AveData_Elec_array(i)) / Log(10) - Log(picmain_ymin) / Log(10)) / (Log(picmain_ymax) / Log(10) - Log(picmain_ymin) / Log(10)) * 5000
                    Pic_Main.Line (X, Pic_Main.Height - 100)-(X, Pic_Main.Height - 100 - Y), &HC000&
                Next i
                If SecNum = 2 Then Data_min = AveData_Elec: Picmain_Ymin = Data_min * 0.9
                
                Pic_Sub.DrawStyle = vbSolid: Pic_Sub.DrawWidth = 1: Pic_Sub.AutoRedraw = True
                If last_Sheath_Flow <> 0 And SecNum >= 3 Then
                    Pic_Sub.Line (60 + (SecNum - 2) * (Pic_Sub.Width - 60) / (Int(Xaxis_range / 15) + 1) / 15, Pic_Sub.Height - (200 + (last_Sheath_Flow - Picsub_YLmin) / (Picsub_YLmax - Picsub_YLmin) * 4000))-(60 + (SecNum - 1) * (Pic_Sub.Width - 60) / (Int(Xaxis_range / 15) + 1) / 15, Pic_Sub.Height - (200 + (Avedata_Sheath_Flow - Picsub_YLmin) / (Picsub_YLmax - Picsub_YLmin) * 4000)), vbMagenta
                End If
                last_Sheath_Flow = Avedata_Sheath_Flow
                
                If last_Aerosol_Flow <> 0 And SecNum >= 3 Then
                    Pic_Sub.Line (60 + (SecNum - 2) * (Pic_Sub.Width - 60) / (Int(Xaxis_range / 15) + 1) / 15, Pic_Sub.Height - (200 + (last_Aerosol_Flow - Picsub_YLmin) / (Picsub_YLmax - Picsub_YLmin) * 4000))-(60 + (SecNum - 1) * (Pic_Sub.Width - 60) / (Int(Xaxis_range / 15) + 1) / 15, Pic_Sub.Height - (200 + (Avedata_Aerosol_Flow - Picsub_YLmin) / (Picsub_YLmax - Picsub_YLmin) * 4000)), vbGreen
                End If
                last_Aerosol_Flow = Avedata_Sheath_Flow
                
                If last_DMA_Volt <> 0 And SecNum >= 3 Then
                    Pic_Sub.Line (60 + (SecNum - 2) * (Pic_Sub.Width - 60) / (Int(Xaxis_range / 15) + 1) / 15, Pic_Sub.Height - (200 + (last_DMA_Volt - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000))-(60 + (SecNum - 1) * (Pic_Sub.Width - 60) / (Int(Xaxis_range / 15) + 1) / 15, Pic_Sub.Height - (200 + (Avedata_DMA_Volt - Picsub_YRmin) / (Picsub_YRmax - Picsub_YRmin) * 4000)), &H80FF&
                End If
                last_DMA_Volt = Avedata_DMA_Volt
                  
              'Next Sample/Sequence End:
                If SecNum = cycletime + 1 Then
                    Txt_samplename.Text = ""
                    Txt_timeleft.Text = ""
                    Txt_timeleft.Text = ""
                  'Sequence END
                    If Cycle_Num = Cycle_times Then
                        Cycle_Num = 0
                        Status = "Standby"
                        List_Sample.Enabled = True
                        Cmd_start.Enabled = True: Cmd_stop.Enabled = False
                        Mnustart.Enabled = True:  Mnustop.Enabled = False:  MnuOpen.Enabled = True
                    Else
                        Status = "Next Sample"
                        Start_time = Start_time + Sample_period / 24 / 60
                    End If
                    Txt_Starttime.Text = Start_time
                End If
            End If
            
          'STANDBY:开机，turn on，after click "stop",sequence completed
          'Next Sample
            If Status = "Standby" Or Status = "Next Sample" Then
              'AO:
                dataScaled_AO(0) = 0: dataScaled_AO(1) = 0
                MnuPS.Enabled = True
                Txt_Elec_data.Text = "Con.: "
                If Avedata_DMA_Volt >= 10 Then
                    Txt_Scan_Volt.Text = "Volt / V:" & VBA.Format(Avedata_DMA_Volt, "#0.00")
                Else
                    Txt_Scan_Volt.Text = "Volt / V:"
                End If
            End If
            
            err = InstantAoCtrl1.WriteChannels(channelStart, channelCount, dataScaled_AO)
            If err <> Success Then Call HandleError(err)
          'Check Flowrates:
'            If (Avedata_Aerosol_Flow >= Aerosol_Flow * 1.1 Or Avedata_Aerosol_Flow <= Aerosol_Flow * 0.9) Or (Avedata_Sheath_Flow >= Sheath_Flow * 1.1 Or Avedata_Sheath_Flow <= Sheath_Flow * 0.9) Then
'                StatusBar1.Panels(3).Text = "Please check aerosol flow!"
'                Txt_status.Text = "Flow Error": Txt_status.BackColor = &HFFFF&
'            Else
                Txt_status.Text = Status: Txt_status.BackColor = &HFFFF&: StatusBar1.Panels(3).Text = Status
'            End If

        End If
        Timermain_time_last = Timermain_time_now
    End If
End Sub
            
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
  
  On Error Resume Next
  
End Sub

Private Sub InitializePortState()
    On Error GoTo 0
    Dim portData As Byte
    portData = 0
    Dim portDir As Byte
    portDir = &HFF
    Dim err As ErrorCode
    err = Success
    Dim portDirs As IPortDirection

    Dim i As Integer
    i = 0
    While (i + StartPort) < InstantDoCtrl1.Features.PortCount And i < PortCountShow
        err = InstantDoCtrl1.ReadPort(i + StartPort, portData)
        If err <> Success Then
            HandleError (err)
            Exit Sub
        End If
        If InstantDoCtrl1.Features.PortProgrammable Then
            portDirs = InstantDoCtrl1.PortDirection
            portDir = portDirs(i).Direction
        End If
        i = i + 1
    Wend
    
End Sub

Private Sub HandleError(ByVal err As ErrorCode)
    Dim utility As BDaqUtility
    Dim errorMessage As String
    Dim res As ErrorCode
        
    Set utility = New BDaqUtility
        
    res = utility.EnumToString("ErrorCode", err, errorMessage)
    
    If err <> Success Then
        MsgBox "Sorry ! There're some errors happened, the error code is: " & errorMessage
    End If
End Sub

