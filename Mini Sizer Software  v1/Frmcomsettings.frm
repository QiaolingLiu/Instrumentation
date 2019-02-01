VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frmsettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties & Settings"
   ClientHeight    =   9630
   ClientLeft      =   2985
   ClientTop       =   390
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   11790
   Begin VB.Frame Frame2 
      Caption         =   " Flow (lpm) "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   3
      Top             =   8580
      Width           =   11600
      Begin VB.TextBox Txt_sheath_flow 
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
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   900
      End
      Begin VB.TextBox Txt_aerosol_flow 
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
         Left            =   1750
         TabIndex        =   4
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "Q_Aerosol :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   7
         Top             =   400
         Width           =   1800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Q_Sheath Flow :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   6
         Top             =   400
         Width           =   1800
      End
   End
   Begin VB.CommandButton Cmd_Setting_exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10440
      TabIndex        =   2
      Top             =   6480
      Width           =   1050
   End
   Begin VB.CommandButton Cmd_Setting_OK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10440
      TabIndex        =   1
      Top             =   4800
      Width           =   1050
   End
   Begin VB.CommandButton Cmd_Setting_reset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10440
      TabIndex        =   0
      Top             =   5640
      Width           =   1050
   End
   Begin VB.Frame Frame17 
      Caption         =   " Voltage Settings "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4100
      Left            =   120
      TabIndex        =   8
      Top             =   30
      Width           =   11600
      Begin VB.OptionButton Opt_volt_P 
         Caption         =   "Positive Voltage"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   32
         Top             =   380
         Width           =   2500
      End
      Begin VB.OptionButton Opt_volt_N 
         Caption         =   "Negative Voltage"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3050
         TabIndex        =   31
         Top             =   380
         Width           =   2500
      End
      Begin VB.Frame Frame8 
         Height          =   3250
         Left            =   5480
         TabIndex        =   13
         Top             =   750
         Width           =   6000
         Begin VB.Frame Frame23 
            Height          =   2650
            Left            =   0
            TabIndex        =   15
            Top             =   600
            Width           =   6000
            Begin VB.TextBox Txt_scantime_U_1 
               Alignment       =   2  'Center
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
               Left            =   2400
               TabIndex        =   24
               Top             =   840
               Width           =   850
            End
            Begin VB.TextBox Txt_size_U 
               Alignment       =   2  'Center
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
               Left            =   3675
               TabIndex        =   23
               Top             =   1425
               Width           =   850
            End
            Begin VB.TextBox Txt_Scan_volt_U 
               Alignment       =   2  'Center
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
               Left            =   3960
               TabIndex        =   22
               Top             =   2050
               Width           =   1200
            End
            Begin VB.TextBox Txt_size_D 
               Alignment       =   2  'Center
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
               Left            =   2400
               TabIndex        =   21
               Top             =   1440
               Width           =   850
            End
            Begin VB.TextBox Txt_Scan_volt_D 
               Alignment       =   2  'Center
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
               Left            =   2400
               TabIndex        =   20
               Top             =   2050
               Width           =   1200
            End
            Begin VB.CommandButton Cmd_maxrange 
               Caption         =   "Max Range"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   4640
               TabIndex        =   19
               Top             =   1425
               Width           =   1200
            End
            Begin VB.OptionButton Opt_up_scan 
               Caption         =   "Up San Mode"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   255
               TabIndex        =   18
               Top             =   240
               Value           =   -1  'True
               Width           =   2000
            End
            Begin VB.OptionButton Opt_upanddown_scan 
               Caption         =   "Up and Down Scan Mode"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   2715
               TabIndex        =   17
               Top             =   240
               Width           =   2955
            End
            Begin VB.TextBox Txt_scantime_D_1 
               Alignment       =   2  'Center
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
               Left            =   4680
               TabIndex        =   16
               Top             =   840
               Width           =   850
            End
            Begin VB.Label Label2 
               Caption         =   "<5 kV"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   5280
               TabIndex        =   58
               Top             =   2160
               Width           =   600
            End
            Begin VB.Label Label23 
               Caption         =   "Scan Time _ Up (s) :"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   255
               TabIndex        =   30
               Top             =   885
               Width           =   2100
            End
            Begin VB.Label Label57 
               Alignment       =   2  'Center
               Caption         =   "~"
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
               Left            =   3315
               TabIndex        =   29
               Top             =   1485
               Width           =   255
            End
            Begin VB.Label Label58 
               Alignment       =   2  'Center
               Caption         =   "~"
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
               Left            =   3650
               TabIndex        =   28
               Top             =   2123
               Width           =   255
            End
            Begin VB.Label Label59 
               Caption         =   "Size Range (nm) :"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   255
               TabIndex        =   27
               Top             =   1460
               Width           =   2100
            End
            Begin VB.Label Label60 
               Caption         =   "Scan Voltage (V) :"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   255
               TabIndex        =   26
               Top             =   2070
               Width           =   2100
            End
            Begin VB.Label Label51 
               Caption         =   "_ Down (s) :"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3400
               TabIndex        =   25
               Top             =   890
               Width           =   1260
            End
         End
         Begin VB.OptionButton Opt_expscanmode 
            Caption         =   "Exp-Scan Mode"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2000
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   2000
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3250
         Left            =   150
         TabIndex        =   9
         Top             =   750
         Width           =   5235
         Begin VB.OptionButton Opt_stepmode 
            Caption         =   "Stepping Mode"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1600
            TabIndex        =   11
            Top             =   150
            Width           =   2035
         End
         Begin VB.TextBox Txt_stepping_set 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   10
            Top             =   2160
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid MSFL_stepping 
            Height          =   2505
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   4419
            _Version        =   393216
            Rows            =   10
            Cols            =   4
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
      End
   End
   Begin VB.Frame Frame19 
      Caption         =   "Sampling Settings"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   120
      TabIndex        =   38
      Top             =   4200
      Width           =   9975
      Begin VB.Frame Frame1 
         Height          =   1050
         Left            =   2040
         TabIndex        =   60
         Top             =   2100
         Width           =   7695
         Begin VB.OptionButton Opt_others 
            Caption         =   "Others"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   65
            Top             =   600
            Width           =   1155
         End
         Begin VB.OptionButton Opt_immediately 
            Caption         =   "Immediately"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   2000
         End
         Begin VB.ComboBox Cmb_others_h 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            ItemData        =   "Frmcomsettings.frx":0000
            Left            =   1485
            List            =   "Frmcomsettings.frx":0002
            TabIndex        =   63
            Text            =   "0"
            Top             =   540
            Width           =   800
         End
         Begin VB.ComboBox Cmb_others_s 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            ItemData        =   "Frmcomsettings.frx":0004
            Left            =   3405
            List            =   "Frmcomsettings.frx":0006
            TabIndex        =   62
            Text            =   "00"
            Top             =   540
            Width           =   800
         End
         Begin VB.ComboBox Cmb_others_m 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            ItemData        =   "Frmcomsettings.frx":0008
            Left            =   2445
            List            =   "Frmcomsettings.frx":000A
            TabIndex        =   61
            Text            =   "00"
            Top             =   540
            Width           =   800
         End
         Begin VB.Label Label76 
            Alignment       =   2  'Center
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2205
            TabIndex        =   67
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label77 
            Alignment       =   2  'Center
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3165
            TabIndex        =   66
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Frame Frame24 
         Height          =   750
         Left            =   2040
         TabIndex        =   44
         Top             =   1360
         Width           =   7695
         Begin VB.TextBox Txt_Sampleperiod_h 
            Alignment       =   2  'Center
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
            Left            =   1800
            TabIndex        =   47
            Top             =   240
            Width           =   750
         End
         Begin VB.TextBox Txt_Sampleperiod_m 
            Alignment       =   2  'Center
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
            Left            =   3000
            TabIndex        =   46
            Top             =   240
            Width           =   750
         End
         Begin VB.TextBox Txt_Cycle_Times 
            Alignment       =   2  'Center
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
            Left            =   6480
            TabIndex        =   45
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   ">1"
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
            Left            =   7320
            TabIndex        =   68
            Top             =   320
            Width           =   255
         End
         Begin VB.Label Label66 
            Caption         =   "Sample period :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   51
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Label73 
            Alignment       =   2  'Center
            Caption         =   "h"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   50
            Top             =   315
            Width           =   255
         End
         Begin VB.Label Label74 
            Alignment       =   2  'Center
            Caption         =   "min"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   49
            Top             =   315
            Width           =   375
         End
         Begin VB.Label Label50 
            Caption         =   "Cycle times :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4920
            TabIndex        =   48
            Top             =   285
            Width           =   1380
         End
      End
      Begin VB.OptionButton Opt_onlyonce 
         Caption         =   "Only Once"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   240
         TabIndex        =   43
         Top             =   1040
         Value           =   -1  'True
         Width           =   1600
      End
      Begin VB.OptionButton Opt_cycle 
         Caption         =   " Cycle"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2040
         TabIndex        =   42
         Top             =   1040
         Width           =   1000
      End
      Begin VB.TextBox Txt_volttime_U 
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
         Left            =   3120
         TabIndex        =   41
         Top             =   475
         Width           =   750
      End
      Begin VB.TextBox Txt_cycletime 
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
         Left            =   8880
         TabIndex        =   40
         Top             =   475
         Width           =   750
      End
      Begin VB.TextBox Txt_volttime_D 
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
         Left            =   6120
         TabIndex        =   39
         Top             =   475
         Width           =   750
      End
      Begin VB.Label Label75 
         Caption         =   "Start Time :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   59
         Top             =   2295
         Width           =   1380
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   55
         Top             =   555
         Width           =   255
      End
      Begin VB.Label Label70 
         Caption         =   "Cycle Time (s):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7200
         TabIndex        =   54
         Top             =   525
         Width           =   1620
      End
      Begin VB.Label lbl_volt_down_time 
         Caption         =   "Volt. Down Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   53
         Top             =   525
         Width           =   1740
      End
      Begin VB.Label Label68 
         Caption         =   "One Cycle :   Volt. Up Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   52
         Top             =   525
         Width           =   2865
      End
   End
   Begin VB.Frame Frame20 
      Caption         =   " Gas / Particle Properties "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   120
      TabIndex        =   33
      Top             =   7500
      Width           =   11600
      Begin VB.TextBox Txt_particle_density 
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
         Left            =   10560
         TabIndex        =   56
         Top             =   400
         Width           =   900
      End
      Begin VB.TextBox Txt_gas_meanfreepath 
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
         Left            =   6480
         TabIndex        =   35
         Top             =   400
         Width           =   1200
      End
      Begin VB.TextBox Txt_gas_viscosity 
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
         Left            =   2400
         TabIndex        =   34
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label82 
         Caption         =   "Particle Density (g/cm3) :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7920
         TabIndex        =   57
         Top             =   450
         Width           =   2820
      End
      Begin VB.Label Label80 
         Caption         =   "Gas Mean Free Path (m) :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3720
         TabIndex        =   37
         Top             =   450
         Width           =   3060
      End
      Begin VB.Label Label79 
         Caption         =   "Gas Viscosity (Pa s) :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   36
         Top             =   450
         Width           =   2220
      End
   End
   Begin VB.Menu Mnusteppingset 
      Caption         =   "Step Setting Menu"
      Visible         =   0   'False
      Begin VB.Menu Mnuadd 
         Caption         =   "Add A New Step"
      End
      Begin VB.Menu Mnudelete 
         Caption         =   "Delete A Step"
      End
      Begin VB.Menu mnusp 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuquit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Frmsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSFL_stepping_col As Integer, MSFL_stepping_row As Integer, MSFL_stepping_row1 As Integer

Private Sampleperiod_h As Integer, Sampleperiod_m As Integer

Private Sub Form_Load()
  
  Dim i As Integer, s() As String
    
    Frmsettings.Top = 0
  
  'Voltage:
    Opt_volt_P.Enabled = True: Opt_volt_N.Enabled = False: Opt_volt_P.Value = True: Opt_volt_N.Value = False
    
    Opt_stepmode.Enabled = False: Opt_stepmode.Value = True: Opt_expscanmode.Value = False
    Opt_expscanmode.Enabled = False: MSFL_stepping.Enabled = False
    Opt_up_scan.Enabled = False: Opt_upanddown_scan.Enabled = False: Opt_up_scan.Value = False: Opt_upanddown_scan.Value = False
    
    MSFL_stepping.Cols = 4: MSFL_stepping.Rows = 7
    MSFL_stepping.WordWrap = True:  MSFL_stepping.HighLight = flexHighlightNever
    MSFL_stepping.ColWidth(0) = 1000
    For i = 1 To MSFL_stepping.Cols - 1
        MSFL_stepping.ColWidth(i) = 1200
    Next i
    For i = 0 To MSFL_stepping.Rows - 1
        MSFL_stepping.RowHeight(i) = 360
        MSFL_stepping.TextMatrix(i, 0) = i
    Next i
    MSFL_stepping.TextMatrix(0, 0) = "Step NO.": MSFL_stepping.TextMatrix(0, 1) = "Dp/nm": MSFL_stepping.TextMatrix(0, 2) = "Voltage/V": MSFL_stepping.TextMatrix(0, 3) = "Time/s"
    Txt_stepping_set.Visible = False
    
    Txt_scantime_U_1.Enabled = False: Txt_scantime_D_1.Enabled = False: Txt_size_D.Enabled = False: Txt_size_U.Enabled = False
    Txt_Scan_volt_D.Enabled = False: Txt_Scan_volt_U.Enabled = False: Cmd_maxrange.Enabled = False
    Txt_scantime_U_1.Text = "": Txt_scantime_D_1.Text = ""
    Txt_size_D.Text = "": Txt_size_U.Text = "": Txt_Scan_volt_D.Text = "": Txt_Scan_volt_U.Text = ""
  
  'Sampling:
    For i = 0 To 23
        Cmb_others_h.AddItem i
    Next i
    For i = 0 To 59
        If i < 10 Then
            Cmb_others_m.AddItem "0" & i
        Else
            Cmb_others_m.AddItem i
        End If
    Next i
    For i = 0 To 59
        If i < 10 Then
            Cmb_others_s.AddItem "0" & i
        Else
            Cmb_others_s.AddItem i
        End If
    Next i
    
    Txt_gas_viscosity.Text = VBA.Format(gas_viscosity, "#.00e+0")
    Txt_gas_meanfreepath.Text = VBA.Format(gas_meanfreepath, "#.00e+0")
    Txt_particle_density.Text = VBA.Format(particle_density, "#.00e+0")
  
    Txt_aerosol_flow.Text = VBA.Format(Aerosol_Flow, "#.00"): Txt_sheath_flow.Text = VBA.Format(Sheath_Flow, "#.00")
    
  'False-New Setting，True-运行中
    If settingflag = False Then
        MSFL_stepping.Enabled = True
        
        Txt_volttime_U.Text = "": Txt_volttime_D.Text = "": Txt_cycletime.Text = ""
        Opt_onlyonce.Enabled = True: Opt_cycle.Enabled = True: Opt_onlyonce.Value = False: Opt_cycle.Value = False
        
        Txt_Sampleperiod_h.Enabled = True: Txt_Sampleperiod_m.Enabled = True: Txt_Cycle_Times.Enabled = True
        Txt_Sampleperiod_h.Text = "": Txt_Sampleperiod_m.Text = "": Txt_Cycle_Times.Text = ""
        Opt_immediately.Enabled = True: Opt_others.Enabled = True: Opt_immediately.Value = False: Opt_others.Value = False
        Cmb_others_h.Enabled = True: Cmb_others_m.Enabled = True: Cmb_others_s.Enabled = True
        Cmb_others_h.Text = "0": Cmb_others_m.Text = "00": Cmb_others_s.Text = "00"
        
        SteppingNum = 0: cycletime = 0: size_Down = 0: size_Up = 0
        Sample_period = 0: Cycle_times = 0: Start_type = "": Start_time = 0: Next_time = 0: End_time = 0
        
        Cmd_Setting_OK.Enabled = True: Cmd_Setting_reset.Enabled = True: Cmd_Setting_exit.Enabled = True
        
     'Main Window - Tables:
        Frmmain.MSFL_data.Clear
        Frmmain.MSFL_data.TextMatrix(0, 0) = "NO.": Frmmain.MSFL_data.TextMatrix(0, 1) = "Size(mid) nm"
        Frmmain.MSFL_data.TextMatrix(0, 2) = "Con #/cm3": Frmmain.MSFL_data.TextMatrix(0, 3) = "Surface um2/cm3"
        Frmmain.MSFL_data.TextMatrix(0, 4) = "Volume um3/cm3": Frmmain.MSFL_data.TextMatrix(0, 5) = "Mass ug/cm3"
        Frmmain.MSFL_data.TextMatrix(0, 6) = "Volt V"
        Frmmain.MSFL_results.Clear
        Frmmain.MSFL_results.TextMatrix(0, 1) = "Number": Frmmain.MSFL_results.TextMatrix(0, 2) = "Surface": Frmmain.MSFL_results.TextMatrix(0, 3) = "Volume"
        Frmmain.MSFL_results.TextMatrix(0, 4) = "Mass"
        Frmmain.MSFL_results.TextMatrix(1, 0) = "Media": Frmmain.MSFL_results.TextMatrix(2, 0) = "Mean": Frmmain.MSFL_results.TextMatrix(3, 0) = "Mode"
        Frmmain.MSFL_results.TextMatrix(4, 0) = "Geo.Mean": Frmmain.MSFL_results.TextMatrix(5, 0) = "Geo.Std. Dev.": Frmmain.MSFL_results.TextMatrix(6, 0) = "Total Con"
      
      'Graphs:
        Xaxis_range = 120: Picmain_Ymax = 0: Picmain_Ymin = 0
        Refresh_PicSub (Xaxis_range)
        Refresh_PicMain (Xaxis_range)
      
      'Initialization:
        Frmmain.Txt_samplename.Text = "": Frmmain.Txt_datafile.Text = ""
        Frmmain.Txt_Starttime.Text = "": Frmmain.Txt_timeleft.Text = ""
        Frmmain.Txt_volt_range_D.Text = "": Frmmain.Txt_volt_range_U.Text = ""
        Frmmain.Txt_size_range_D.Text = "": Frmmain.Txt_size_range_U.Text = ""
        Frmmain.List_Sample.Clear
        Frmmain.List_Sample.Enabled = True
    
    Else
        MSFL_stepping.Enabled = True
        
        Opt_onlyonce.Enabled = True: Opt_cycle.Enabled = True
        Txt_Sampleperiod_h.Enabled = True: Txt_Sampleperiod_m.Enabled = True: Txt_Cycle_Times.Enabled = True
        Opt_immediately.Enabled = True: Opt_others.Enabled = True: Cmb_others_h.Enabled = True: Cmb_others_m.Enabled = True: Cmb_others_s.Enabled = True
        
        If UBound(Stepping()) > 0 Then
            For i = 1 To UBound(Stepping())
                MSFL_stepping.TextMatrix(i, 0) = i
                MSFL_stepping.TextMatrix(i, 1) = Stepping(i).size
                MSFL_stepping.TextMatrix(i, 2) = Stepping(i).Voltage
                MSFL_stepping.TextMatrix(i, 3) = Stepping(i).t
            Next i
        End If
        Txt_volttime_U.Text = "": Txt_volttime_D.Text = "": Txt_cycletime.Text = cycletime
                
        If Cycle_times = 1 Then
            Opt_onlyonce.Value = True: Opt_cycle.Value = False
        Else
            Opt_onlyonce.Value = False: Opt_cycle.Value = True
            If Sample_period > 0 Then Txt_Sampleperiod_h.Text = Int(Sample_period / 60): Txt_Sampleperiod_m.Text = Int(Sample_period Mod 60)
            If Cycle_times > 1 Then Txt_Cycle_Times.Text = Cycle_times
            Select Case Start_type
                Case "Immediately"
                    Opt_immediately.Value = True: Opt_others.Value = False
                Case "Others"
                    Opt_immediately.Value = False: Opt_others.Value = True
                    s() = Split(Start_time, " ")
                    s() = Split(s(1), ":")
                    Cmb_others_h.Text = s(0): Cmb_others_m.Text = s(1): Cmb_others_m.Text = s(2)
            End Select
        End If
               
        Cmd_Setting_OK.Enabled = True: Cmd_Setting_reset.Enabled = True: Cmd_Setting_exit.Enabled = True
    End If
             
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Opt_stepmode_Click()
    
  Dim i As Integer
    
    Opt_stepmode.Value = True
    MSFL_stepping.Enabled = True: Txt_stepping_set.Visible = False: MSFL_stepping_row = 0: MSFL_stepping_col = 0: MSFL_stepping_row1 = 0
    MSFL_stepping.Clear
    MSFL_stepping.TextMatrix(0, 0) = "Step NO.": MSFL_stepping.TextMatrix(0, 1) = "Dp/nm": MSFL_stepping.TextMatrix(0, 2) = "Voltage/V": MSFL_stepping.TextMatrix(0, 3) = "Time/s"
    For i = 1 To MSFL_stepping.Rows - 1
        MSFL_stepping.TextMatrix(i, 0) = i
    Next i
    
End Sub


Private Sub MSFL_stepping_DblClick()
     
    MSFL_stepping_col = MSFL_stepping.MouseCol
    MSFL_stepping_row = MSFL_stepping.MouseRow
    If MSFL_stepping_col > 0 And MSFL_stepping_row > 0 Then
        Txt_stepping_set.Left = MSFL_stepping.Left + MSFL_stepping.ColPos(MSFL_stepping_col) + 50
        Txt_stepping_set.Top = MSFL_stepping.Top + MSFL_stepping.RowPos(MSFL_stepping_row) + 50
        Txt_stepping_set.Width = MSFL_stepping.ColWidth(MSFL_stepping_col)
        Txt_stepping_set.Text = MSFL_stepping.Text
        Txt_stepping_set.Visible = True
        Txt_stepping_set.SetFocus
    Else
        Txt_stepping_set.Visible = False: Txt_stepping_set.Text = ""
        MSFL_stepping_col = 0
        MSFL_stepping_row = 0
    End If
End Sub

Private Sub MSFL_stepping_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MSFL_stepping_row1 = MSFL_stepping.MouseRow
    If Button = 2 And X <= MSFL_stepping.CellWidth And Y > 360 Then
        Frmsettings.PopupMenu Mnusteppingset, 0, X + Frame17.Left + Frame4.Left + MSFL_stepping.Left, _
        Y + Frame17.Top + Frame4.Top + MSFL_stepping.Top
    End If
End Sub

Private Sub MSFL_stepping_Scroll()

    If (MSFL_stepping.Top + MSFL_stepping.RowPos(MSFL_stepping_row)) >= MSFL_stepping.Top + 360 _
        And (MSFL_stepping.Top + MSFL_stepping.RowPos(MSFL_stepping_row)) <= MSFL_stepping.Top + 360 * 6 Then
        Txt_stepping_set.Left = MSFL_stepping.Left + MSFL_stepping.ColPos(MSFL_stepping_col) + 50
        Txt_stepping_set.Top = MSFL_stepping.Top + MSFL_stepping.RowPos(MSFL_stepping_row) + 50
        Txt_stepping_set.Width = MSFL_stepping.ColWidth(MSFL_stepping_col)
        Txt_stepping_set.Visible = True
    Else
        Txt_stepping_set.Visible = False
    End If
    
End Sub

Private Sub Txt_stepping_set_KeyPress(KeyAscii As Integer)
    
  Dim Dp As Single, V As Single
  
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        Exit Sub
    Else
        If KeyAscii = vbKeyReturn Then
            If Txt_stepping_set.Text <> "" Then
                MSFL_stepping.TextMatrix(MSFL_stepping_row, MSFL_stepping_col) = Txt_stepping_set.Text
                If MSFL_stepping_col = 1 Then
                    Dp = Val(Txt_stepping_set.Text)
                    If Dp >= size_min And Dp <= size_max Then
                        V = Dp_to_Volt(Dp)
                        MSFL_stepping.TextMatrix(MSFL_stepping_row, MSFL_stepping_col + 1) = VBA.Format(V, "#.000")
                    Else
                        MsgBox "Input size must be between " & size_min & "nm and " & size_max & "nm !", 48, "Properties & Settings"
                        MSFL_stepping.TextMatrix(MSFL_stepping_row, MSFL_stepping_col) = ""
                        MSFL_stepping.TextMatrix(MSFL_stepping_row, MSFL_stepping_col + 1) = ""
                    End If
                End If
                If MSFL_stepping_col = 2 Then
                    V = Val(Txt_stepping_set.Text)
                    If V <= Volt_max And V >= Volt_min Then
                        Dp = Volt_to_Dp(V)
                        MSFL_stepping.TextMatrix(MSFL_stepping_row, MSFL_stepping_col - 1) = VBA.Format(Dp, "#.000")
                    Else
                        MsgBox "Input voltage must be between 0 and 5000V !", 48, "Properties & Settings"
                        MSFL_stepping.TextMatrix(MSFL_stepping_row, MSFL_stepping_col) = ""
                        MSFL_stepping.TextMatrix(MSFL_stepping_row, MSFL_stepping_col - 1) = ""
                    End If
                End If
            End If
            Txt_stepping_set.Text = ""
            If MSFL_stepping_row < MSFL_stepping.Rows - 1 Then
                MSFL_stepping_row = MSFL_stepping_row + 1
                Txt_stepping_set.Left = MSFL_stepping.Left + MSFL_stepping.ColPos(MSFL_stepping_col) + 50
                Txt_stepping_set.Top = MSFL_stepping.Top + MSFL_stepping.RowPos(MSFL_stepping_row) + 50
                Txt_stepping_set.Width = MSFL_stepping.ColWidth(MSFL_stepping_col)
                Txt_stepping_set.Text = MSFL_stepping.TextMatrix(MSFL_stepping_row, MSFL_stepping_col)
                Txt_stepping_set.SetFocus
                Txt_stepping_set.Visible = True
            Else
                If MSFL_stepping_row = MSFL_stepping.Rows - 1 Then Txt_stepping_set.Visible = False
                MSFL_stepping_col = 0
                MSFL_stepping_row = 0
            End If
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_Cycle_Times_KeyUp(KeyCode As Integer, Shift As Integer)

    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 8 Or KeyCode = 190) Then            '0-9,delete,.
        Txt_Cycle_Times.Text = ""
        Txt_Cycle_Times.SelStart = Len(Txt_Cycle_Times.Text)
    End If
    
End Sub

Private Sub Opt_onlyonce_Click()
    
    Txt_Sampleperiod_h.Enabled = False: Txt_Sampleperiod_m.Enabled = False: Txt_Cycle_Times.Enabled = False
    Txt_Sampleperiod_h.Text = "": Txt_Sampleperiod_m.Text = "": Txt_Cycle_Times.Text = ""
    
 Dim i As Integer
    cycletime = 0
    For i = 1 To MSFL_stepping.Rows - 1
        cycletime = cycletime + VBA.Val(MSFL_stepping.TextMatrix(i, 3))
    Next i
    Txt_cycletime.Text = cycletime

End Sub

Private Sub Opt_cycle_Click()

    Txt_Sampleperiod_h.Enabled = True: Txt_Sampleperiod_m.Enabled = True: Txt_Cycle_Times.Enabled = True
   
   Dim i As Integer
    cycletime = 0
    For i = 1 To MSFL_stepping.Rows - 1
        cycletime = cycletime + VBA.Val(MSFL_stepping.TextMatrix(i, 3))
    Next i
    Txt_cycletime.Text = cycletime
    
End Sub

Private Sub Txt_Sampleperiod_h_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 8 Or KeyCode = 190) Then
        Txt_Sampleperiod_h.Text = ""
        Txt_Sampleperiod_h.SelStart = Len(Txt_Sampleperiod_h.Text)
    End If
End Sub

Private Sub Txt_Sampleperiod_m_KeyUp(KeyCode As Integer, Shift As Integer)
    If Val(Txt_Sampleperiod_m.Text) >= 60 Then
        Txt_Sampleperiod_m.Text = ""
        Txt_Sampleperiod_m.SelStart = Len(Txt_Sampleperiod_m)
    End If
    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 8 Or KeyCode = 190) Then
        Txt_Sampleperiod_m.Text = ""
        Txt_Sampleperiod_m.SelStart = Len(Txt_Sampleperiod_m.Text)
    End If
End Sub

Private Sub Opt_others_Click()
    Cmb_others_h.Enabled = True: Cmb_others_m.Enabled = True: Cmb_others_s.Enabled = True
End Sub

Private Sub cmd_setting_ok_Click()
  
  Dim i As Integer, myfilename$
  Dim stepping_size_d As Single, stepping_size_u As Single, stepping_volt_d As Single, stepping_volt_u As Single
  Dim setting_complete1 As Boolean, setting_complete2 As Boolean, setting_complete3 As Boolean
  
    setting_complete1 = False: setting_complete2 = False: setting_complete3 = False
    
    If Opt_stepmode.Value = True Then
        SteppingNum = 0
        For i = 1 To MSFL_stepping.Rows - 1
            If MSFL_stepping.TextMatrix(i, 1) <> "" And MSFL_stepping.TextMatrix(i, 3) <> "" Then SteppingNum = SteppingNum + 1
        Next i
        If SteppingNum > 0 Then
            setting_complete1 = True
            ReDim Stepping(1 To SteppingNum) As Steppingset
            For i = 1 To SteppingNum
                Stepping(i).size = Val(MSFL_stepping.TextMatrix(i, 1))
                Stepping(i).Voltage = Val(MSFL_stepping.TextMatrix(i, 2))
                Stepping(i).t = Val(MSFL_stepping.TextMatrix(i, 3))
                If i = 1 Then
                    Stepping(i).tacu = Stepping(i).t
                    stepping_size_d = Stepping(i).size: stepping_size_u = Stepping(i).size
                    stepping_volt_d = Stepping(i).Voltage: stepping_volt_u = Stepping(i).Voltage
                Else
                    Stepping(i).tacu = Stepping(i - 1).tacu + Stepping(i).t
                    If stepping_size_u < Stepping(i).size Then stepping_size_u = Stepping(i).size: stepping_volt_u = Stepping(i).Voltage
                    If stepping_size_d > Stepping(i).size Then stepping_size_d = Stepping(i).size: stepping_volt_d = Stepping(i).Voltage
                End If
            Next i
            cycletime = Stepping(SteppingNum).tacu
            Xaxis_range = Stepping(SteppingNum).tacu + 5
        End If
    End If
        
    If Opt_onlyonce.Value = True Then
        setting_complete2 = True
        Cycle_times = 1
        Start_type = ""
    End If
    If Opt_cycle.Value = True Then
        If (Txt_Sampleperiod_h.Text <> "" Or Txt_Sampleperiod_m.Text <> "") And Val(Txt_Cycle_Times.Text) > 1 And (Opt_immediately.Value = True Or (Opt_others.Value = True)) Then
            setting_complete2 = True
            Sample_period = Val(Txt_Sampleperiod_h.Text) * 60 + Val(Txt_Sampleperiod_m.Text)      'min
            Cycle_times = Val(Txt_Cycle_Times.Text)
        End If
    End If
       
    If Opt_immediately.Value = True Then Start_type = "Immediately": setting_complete3 = True
    If Opt_others.Value = True Then
        Start_type = "Others"
        Start_time = VBA.Date + Val(Cmb_others_h.Text) / 24 + Val(Cmb_others_m.Text) / 24 / 60 + Val(Cmb_others_s.Text) / 24 / 60 / 60 'mm-dd-yyyy hh:mm:ss
        setting_complete3 = True
        If Opt_cycle.Value = True Then
            Next_time = Start_time + Sample_period / 24 / 60
            End_time = Start_time + Sample_period / 24 / 60 * (Cycle_times - 1) + cycletime / 24 / 60 / 60
        End If
    End If
   
    If setting_complete1 = True And setting_complete2 = True And setting_complete3 = True Then
      
        DoEvents
        On Error Resume Next
        Frmmain.CommonDialog1.CancelError = True
        Frmmain.CommonDialog1.InitDir = App.Path & "\Mini-Sizer Data\Raw data"
        Frmmain.CommonDialog1.DialogTitle = "New file / *.TXT"
        i = 1
        myfilename = "Rawdata_" & Month(VBA.Date) & "-" & Day(VBA.Date) & "-" & Year(VBA.Date) & "_" & i & ".TXT"
        Do While fso.FileExists(App.Path & "\Mini-Sizer Data\Raw data\" & myfilename)
            i = i + 1
            myfilename = "Rawdata_" & Month(VBA.Date) & "-" & Day(VBA.Date) & "-" & Year(VBA.Date) & "_" & i & ".TXT"
        Loop
        Frmmain.CommonDialog1.FileName = myfilename
        Frmmain.CommonDialog1.Filter = "Document(*.TXT)|*.TXT"
        Frmmain.CommonDialog1.FilterIndex = 1
        Frmmain.CommonDialog1.Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNOverwritePrompt
        Frmmain.CommonDialog1.ShowSave                                               '"保存"对话框
        If err = cdlCancel Then GoTo 1
        Datafilepath = Frmmain.CommonDialog1.FileName
        Frmmain.Txt_datafile = Datafilepath                                          '显示保存文件路径
      'Saving:
        Open Datafilepath For Output As #1
        Print #1, VBA.Date & VBA.Time & "  /  DATA"
        Print #1, "DMA Length (cm) :   " & DMA_Length
        Print #1, "DMA Width (cm) :   " & DMA_Width
        Print #1, "DMA Height (cm) :   " & DMA_Height
        Print #1, "Voltage Mode :   " & "Positive" & ", " & "Stepping"
        If SteppingNum > 0 Then
            Print #1, Tab(1); "Steps:" & SteppingNum; Tab(11); " Size"; Tab(21); "Voltage"; Tab(31); "Time"
            For i = 1 To SteppingNum
                Print #1, Tab(1); "NO: " & i; Tab(10); Stepping(i).size; Tab(20); Stepping(i).Voltage; Tab(30); Stepping(i).t
            Next i
        End If
        Print #1, "Gas viscosity (Pa s) :   " & Txt_gas_viscosity.Text
        Print #1, "Gas mean free path (m) :   " & Txt_gas_meanfreepath.Text
        Print #1, "Particle Density (g/cm3) :   " & particle_density
        Close #1
1:
        If Start_time <> 0 Then
            Frmmain.Txt_Starttime = Start_time
        Else
            Frmmain.Txt_Starttime = "Immediately"
        End If
        Sample_Num = 0: Frmmain.Txt_samplename = "Sample 1"
        Frmmain.Txt_size_range_D = VBA.Format(stepping_size_d, "#.00"): Frmmain.Txt_size_range_U = VBA.Format(stepping_size_u, "#.00")
        Frmmain.Txt_volt_range_D = VBA.Format(stepping_volt_d, "#.00"): Frmmain.Txt_volt_range_U = VBA.Format(stepping_volt_u, "#.00")
       
       'Refresh Pics
        Picmain_Ymax = 1000: Picmain_Ymin = 0

        Refresh_PicSub (Xaxis_range)
        Refresh_PicMain (Xaxis_range)
    
        settingflag = True
        Cycle_Num = 0
        If Start_type = "Immediately" Then
            Frmmain.Cmd_start.Enabled = True
        Else
            Frmmain.Cmd_start.Enabled = False
        End If
        Frmmain.Cmd_stop.Enabled = True
        Sleep (500)
        Unload Me
    Else
        MsgBox "Missing! Please check all the settings !", 48, "Properties & Settings"
    End If

End Sub
  
Private Sub Cmd_Setting_reset_Click()
    
    MSFL_stepping.Clear
    MSFL_stepping.TextMatrix(0, 0) = "Step NO.": MSFL_stepping.TextMatrix(0, 1) = "Dp/nm": MSFL_stepping.TextMatrix(0, 2) = "Voltage/V": MSFL_stepping.TextMatrix(0, 3) = "Time/s"
    Txt_stepping_set.Visible = False
           
    Txt_cycletime.Text = ""
    Opt_onlyonce.Value = False: Opt_cycle.Value = False
    Txt_Sampleperiod_h.Enabled = False: Txt_Sampleperiod_m.Enabled = False: Txt_Cycle_Times.Enabled = False
    Txt_Sampleperiod_h.Text = "": Txt_Sampleperiod_m.Text = "": Txt_Cycle_Times.Text = ""
    Opt_immediately.Enabled = False: Opt_others.Enabled = False
    Cmb_others_h.Enabled = False: Cmb_others_m.Enabled = False: Cmb_others_s.Enabled = False
    Cmb_others_h.Text = "0": Cmb_others_m.Text = "00": Cmb_others_s.Text = "00"
    
    SteppingNum = 0:  cycletime = 0:
    Sample_period = 0: Cycle_times = 0: Start_type = "": Start_time = 0: Next_time = 0: End_time = 0
       
End Sub

Private Sub Cmd_Setting_exit_Click()
    Unload Me
End Sub


Private Sub mnuadd_Click()
  
  Dim i As Integer
    
    MSFL_stepping.AddItem (MSFL_stepping.Rows)
    MSFL_stepping.RowHeight(MSFL_stepping.Rows - 1) = 360
    i = MSFL_stepping.Rows - 1
    Do While Not i = MSFL_stepping_row1
        MSFL_stepping.TextMatrix(i, 0) = i
        MSFL_stepping.TextMatrix(i, 1) = MSFL_stepping.TextMatrix(i - 1, 1)
        MSFL_stepping.TextMatrix(i, 2) = MSFL_stepping.TextMatrix(i - 1, 2)
        i = i - 1
    Loop
    MSFL_stepping.TextMatrix(i, 1) = ""
    MSFL_stepping.TextMatrix(i, 2) = ""
    
End Sub
Private Sub mnudelete_Click()

  Dim i As Integer
    MSFL_stepping.Rows = MSFL_stepping.Rows - 1
    i = MSFL_stepping_row1
    If i < MSFL_stepping.Rows Then
        Do While Not i = MSFL_stepping.Rows - 1
            MSFL_stepping.TextMatrix(i, 0) = i
            MSFL_stepping.TextMatrix(i, 1) = MSFL_stepping.TextMatrix(i + 1, 1)
            MSFL_stepping.TextMatrix(i, 2) = MSFL_stepping.TextMatrix(i + 1, 2)
            i = i + 1
        Loop
        MSFL_stepping.TextMatrix(i, 1) = ""
        MSFL_stepping.TextMatrix(i, 2) = ""
    End If
    
End Sub



