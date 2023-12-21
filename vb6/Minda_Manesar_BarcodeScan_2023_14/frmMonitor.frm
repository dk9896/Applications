VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMonitor 
   Caption         =   "Minda_Manesar_BarcodeScan_2023_14"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   15630
   Begin VB.Timer Timer9 
      Interval        =   1000
      Left            =   7200
      Top             =   4080
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.CommandButton Command3 
         Caption         =   "GO Command"
         Height          =   855
         Left            =   120
         TabIndex        =   60
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Timer Timer10 
         Interval        =   2000
         Left            =   3960
         Top             =   5400
      End
      Begin VB.TextBox txtScanBarcode 
         Height          =   615
         Left            =   960
         TabIndex        =   56
         Top             =   2400
         Width           =   10455
      End
      Begin VB.Frame frmCoupler 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Index           =   0
         Left            =   7320
         TabIndex        =   49
         Top             =   3360
         Width           =   4095
         Begin VB.TextBox txtProductionCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   20.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   690
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3000
            Width           =   2790
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00008080&
            Caption         =   "RST"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtNGCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   18
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1080
            Width           =   2310
         End
         Begin VB.TextBox txtOKCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   18
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   240
            Width           =   2310
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Production Counter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1200
            TabIndex        =   59
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "OK Count"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "NG Count"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1320
            Width           =   855
         End
      End
      Begin VB.TextBox txtModelDesc 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1020
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "MODEL DESC"
         Top             =   240
         Width           =   10215
      End
      Begin VB.TextBox txtCommandLine 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   44
         Text            =   "frmMonitor.frx":0000
         Top             =   8160
         Width           =   11175
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10320
         TabIndex        =   39
         Top             =   360
         Width           =   2535
         Begin VB.TextBox txtCycleTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   360
            Width           =   720
         End
         Begin VB.Shape ShapePLCState 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   1320
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Cycle Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "sec"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   41
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label17 
            Caption         =   "PLC Comm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame FrmResult 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   960
         TabIndex        =   36
         Top             =   3360
         Width           =   4455
         Begin VB.Label lblNg 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   140.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   3195
            Left            =   60
            TabIndex        =   38
            Top             =   120
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.Label lblGo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   140.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   3195
            Left            =   30
            TabIndex        =   37
            Top             =   120
            Visible         =   0   'False
            Width           =   4365
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   11640
         TabIndex        =   34
         Top             =   8160
         Width           =   1335
         Begin VB.CommandButton CmdClose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   0
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMonitor.frx":0012
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Breakdown"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox PictureBreakdown 
         BackColor       =   &H80000010&
         Height          =   6015
         Left            =   2040
         ScaleHeight     =   5955
         ScaleWidth      =   8835
         TabIndex        =   26
         Top             =   9600
         Visible         =   0   'False
         Width           =   8895
         Begin VB.CommandButton cmdrunningbreakdown 
            BackColor       =   &H000080FF&
            Caption         =   "Running Breakdown"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmdfullbreakdown 
            BackColor       =   &H000000FF&
            Caption         =   "Full Breakdown"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmdgolive 
            BackColor       =   &H0000FF00&
            Caption         =   "Go Live"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox txtbreakdownsummary 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   2280
            TabIndex        =   28
            Top             =   4440
            Width           =   4575
         End
         Begin VB.CommandButton cmdclosebreakdownscreen 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   7200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMonitor.frx":0C54
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   4680
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "BreakDown Summary"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   32
            Top             =   4800
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         Height          =   4455
         Left            =   10680
         ScaleHeight     =   4395
         ScaleWidth      =   7155
         TabIndex        =   1
         Top             =   9360
         Visible         =   0   'False
         Width           =   7215
         Begin VB.TextBox txtDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1800
            TabIndex        =   18
            Text            =   "0000"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtOnline 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   3120
            TabIndex        =   17
            Text            =   "0000"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtCalibrateValue 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   4440
            TabIndex        =   16
            Text            =   "0000"
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdSetCalibration 
            Caption         =   "Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   5760
            TabIndex        =   15
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1800
            TabIndex        =   14
            Text            =   "0000"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtOnline 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   3120
            TabIndex        =   13
            Text            =   "0000"
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtCalibrateValue 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   4440
            TabIndex        =   12
            Text            =   "0000"
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton cmdSetCalibration 
            Caption         =   "Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   5760
            TabIndex        =   11
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   1800
            TabIndex        =   10
            Text            =   "0000"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtOnline 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   3120
            TabIndex        =   9
            Text            =   "0000"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtCalibrateValue 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   4440
            TabIndex        =   8
            Text            =   "0000"
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton cmdSetCalibration 
            Caption         =   "Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   5760
            TabIndex        =   7
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   1800
            TabIndex        =   6
            Text            =   "0000"
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtOnline 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   3120
            TabIndex        =   5
            Text            =   "0000"
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txtCalibrateValue 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   4440
            TabIndex        =   4
            Text            =   "0000"
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton cmdSetCalibration 
            Caption         =   "Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   5760
            TabIndex        =   3
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton cmdCloseLoadCalibration 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   5400
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMonitor.frx":1896
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   3480
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Display"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1920
            TabIndex        =   25
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Online"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   24
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   4560
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Load Cell - 4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   195
            TabIndex        =   22
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Load Cell - 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   195
            TabIndex        =   21
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Load Cell - 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   195
            TabIndex        =   20
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Load Cell - 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   195
            TabIndex        =   19
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   12360
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.Timer Timer5 
            Left            =   2640
            Top             =   1080
         End
         Begin VB.TextBox txtServoSpeedSet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   360
            Width           =   1440
         End
         Begin VB.Timer Timer4 
            Left            =   1320
            Top             =   960
         End
         Begin VB.Timer Timer2 
            Left            =   480
            Top             =   960
         End
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   120
            Top             =   960
         End
         Begin VB.Timer Timer3 
            Left            =   840
            Top             =   960
         End
         Begin VB.Timer Timer6 
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer11 
            Enabled         =   0   'False
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer12 
            Enabled         =   0   'False
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer13 
            Enabled         =   0   'False
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer7 
            Left            =   360
            Top             =   1320
         End
         Begin VB.Timer Timer8 
            Interval        =   60000
            Left            =   840
            Top             =   240
         End
         Begin MSWinsockLib.Winsock WinSock1 
            Left            =   1920
            Top             =   960
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin MSCommLib.MSComm MSComm1 
            Left            =   120
            Top             =   240
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
         Begin TextPrinter.JustPrinter JustPrinter1 
            Height          =   495
            Left            =   1080
            TabIndex        =   47
            Top             =   240
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
      End
      Begin VB.Label Label15 
         Caption         =   "Scan Barcode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   5160
         TabIndex        =   57
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Image ImgPart 
         Height          =   1815
         Left            =   17400
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "OK Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   50
         Top             =   9480
         Width           =   855
      End
      Begin VB.Shape shpSwContact4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   4560
         Top             =   9480
         Width           =   615
      End
   End
   Begin VB.Timer Timer14 
      Left            =   9480
      Top             =   5280
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   -1800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim MsgCode As Integer
Dim Pulse As Boolean
Dim pulse1 As Boolean
Dim pulse2 As Boolean
Dim pulse3 As Boolean
Dim pulse4 As Boolean
Dim PulseScan As Boolean
Dim pulseBreakdown As Boolean
Dim PulseReset As Boolean
Dim pulsePrinterBypass As Boolean
Dim FSO As New FileSystemObject
Dim ExcelFileName As String
Dim Row As Long
Dim Col As Long
Dim setCouplerCounter As Integer
Dim setBatchCounter As Integer
Dim IP_HOST As String
Dim IP_PORT As String
Dim PartCode As String
'----------------
Dim PLC_Communication_Error As Boolean
Dim MsgText() As String
Dim MsgColor() As Integer
Dim MsgCount As Integer
Dim CloseScreen As Boolean
Dim runningreportdate As Date
Dim runningreportshift As String
Dim ModelNo As Integer
Dim ScanValidationBypass As Boolean
Dim rsdb As ADODB.Recordset

Private Sub cmdClose_Click()
CloseScreen = True
CloseMe
End Sub

Private Sub CloseMe()

If MSComm1.PortOpen = True Then MSComm1.PortOpen = False

frmmenu.Show
Unload Me

End Sub

Private Sub CmdNgCounter_Click()
  If MsgBox("Are you Sure You Want To Reset NG Counter", vbInformation + vbYesNo) = vbYes Then
    'txtNGCounter.Text = 0
    'SaveCounterValue
  End If
End Sub

Private Sub CmdOKCounter_Click()
If MsgBox("Are you Sure You Want To Reset OK Counter", vbInformation + vbYesNo) = vbYes Then
    'txtOKCounter.Text = 0
    'SaveCounterValue
  End If
End Sub


Private Sub Command1_Click()
If MsgBox("Do you want to reset counter", vbYesNo) = vbYes Then
  SaveSetting App.Title, ModelName, "OKCounter", 0
  SaveSetting App.Title, ModelName, "NGCounter", 0
  txtOKCounter.Text = 0
  txtNGCounter.Text = 0
End If
End Sub

Private Sub Command3_Click()
      pulse1 = False
      lblGo.Visible = True
        GetCounterValue
        txtOKCounter.Text = Val(txtOKCounter.Text) + 1
        txtProductionCounter.Text = Val(txtProductionCounter.Text) + 1
        SaveReport 1, barcode
        Print_Click
        SaveCounterValue

End Sub

Private Sub Form_Unload(Cancel As Integer)
If CloseScreen = False Then
    CloseMe
Else
    CloseScreen = False
End If
End Sub

Public Sub ConnectToPLC()
On Error GoTo Error
Dim Sql As String
Dim Rs As ADODB.Recordset

   'To Load Com port in Monitor
   Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   Dim ComPort(3) As Integer
   Dim ComPortBP(3) As Integer
   ''ComPort(1) = rs("ComPort1")
''    ComPort(2) = Rs("ComPort2")
    ''ComPortBP(1) = rs("ComPortBP1")
''      ComPortBP(2) = Rs("ComPortBP2")
   PrinterName = Rs("PrinterName1")
   Initialise
   Winsock1.Protocol = sckTCPProtocol
   'txtIP.Text = WinSock1.LocalIP
   IP_HOST = Rs("PLC_IP") '"192.168.1.30"
   IP_PORT = Rs("PLC_Port")
Exit Sub
Error:
If Err.Number = 8002 Then
    MsgBox "Com Port " & ComPort(Erl) & " Not Working", vbInformation
ElseIf Err.Number = 8005 Then
    MsgBox "Com Port " & ComPort(Erl) & " Already Open", vbInformation
Else
    MsgBox Error, vbInformation
End If
End Sub

Private Sub Form_Load()
''On Error GoTo Error
Me.WindowState = 2
UserAccess
Frame1.Top = ((Screen.Height - Frame1.Height) / 2) - 500
Frame1.Left = ((Screen.Width - Frame1.Width) / 2)
LoadSettingsData
Call Load_Message_File
'SaveReport 1, "aadsvvd"
LoadGrid
PLcdata(340) = 1
GetCounterValue
ConnectToPLC
'txtScanBarcode.SetFocus
Timer1.Enabled = True
Timer1.Interval = 1000
Timer2.Enabled = True
Timer2.Interval = 1000
Timer3.Interval = 500
Timer3.Enabled = True

'txtDate.Text = Date
'txttime.Text = Format(Time(), "hh:mm:ss")
'txtOperName.Text = LoginUser
Pulse = False
Exit Sub
End Sub

Private Sub UserAccess()
   If AccessType = "0" Then 'Disable or Hide For Operator
   ElseIf AccessType = "1" Then 'Disable or Hide for AccessType 1
   ElseIf AccessType = "2" Then 'Show All Which Will Disable or Hide For One
   End If
End Sub

Private Function AssignPLCdata()
On Error GoTo Error
   MsgCode = PLcdata(2008)
   
   If PLcdata(2002) = 0 And pulse2 = False Then
        pulse2 = True
        PLcdata(2102) = 0
        txtScanBarcode.Text = ""
        txtScanBarcode.BackColor = vbWhite
        txtScanBarcode.Locked = True
   ElseIf PLcdata(2002) = 1 And pulse2 = True Then
        pulse2 = False
        txtScanBarcode.Text = ""
        txtScanBarcode.BackColor = vbWhite
        txtScanBarcode.Locked = False
        txtScanBarcode.SetFocus
   End If

   
   
   'ShapeColorfunction PLcdata(110), &H1, &H2, txtProcessName
   txtCycleTime.Text = Format(PLcdata(2007) / 10, "0.0")
   If PLcdata(2009) = 0 And pulse1 = False Then
      pulse1 = True
      lblGo.Visible = False
      lblNg.Visible = False
   ElseIf PLcdata(2009) = 1 And pulse1 = True Then
      pulse1 = False
      lblGo.Visible = True
        GetCounterValue
        txtOKCounter.Text = Val(txtOKCounter.Text) + 1
        txtProductionCounter.Text = Val(txtProductionCounter.Text) + 1
        Print_Click
        SaveReport 1, barcode
        
        SaveCounterValue
        
   ElseIf PLcdata(2009) = 2 And pulse1 = True Then
      pulse1 = False
      GetCounterValue
      lblNg.Visible = True
      txtNGCounter.Text = Val(txtNGCounter.Text) + 1
      SaveReport 2, ""
      SaveCounterValue
   End If
      
Exit Function
Error:
   ErrorLog Err.Number, Err.Description & "---", Erl, Me.Name, "Assign PLC Data"
   Resume Next
End Function
Private Sub saveToNotepad()
On Error GoTo Error
Dim tempdate, tempCounter As String

tempdate = GetDateCode
MarkDateCode = tempdate
tempCounter = Format(txtProductionCounter.Text, "00000")
MarkSerialNo = tempCounter
Dim markingFileData As String
markingFileData = NotepadModelName & vbCrLf & CustomerDrawingNo & vbCrLf

Dim FSO As New FileSystemObject
Dim iFile As String
Dim iFileNo As Integer

If FSO.FolderExists(NotePadPath) = True Then
    iFile = NotePadPath & "\PrintFile.txt"
    If FSO.FileExists(iFile) = True Then
        FSO.DeleteFile iFile, True
        
    End If
    FSO.CreateTextFile iFile
    iFileNo = FreeFile
    Open iFile For Append As iFileNo
    Print #iFileNo, NotepadModelName
    Close iFileNo
    Open iFile For Append As iFileNo
    Print #iFileNo, CustomerDrawingNo
    Close iFileNo
    Open iFile For Append As iFileNo
    Print #iFileNo, tempdate
    Close iFileNo
    Open iFile For Append As iFileNo
    Print #iFileNo, tempCounter
    Close iFileNo
Else
    ErrorLog 0, "Notepad FolderPath not found", "", "", ""
End If
Exit Sub
Error:
  ErrorLog Err.Number, Err.Description, Err.Source, "", ""
  Err.Clear
End Sub
Private Function GetDateCode() As String
Dim month
Dim day
Dim year
Dim Swe() As String
month = Format(Now(), "MM")
day = Format(Now(), "dd")

'If day > 9 Then
'A = day - 9
'X = "0,A,B,C,D,E,F,G,H,I,J,K"
'Swe = Split(X, ",")
'day = Swe(A)



year = Mid(Format(Now(), "yy"), 2, 1)
If Val(day) < 10 Then
    day = Val(day)
ElseIf Val(day) = 10 Then
    day = 0
ElseIf Val(day) = 11 Then
    day = "A"
ElseIf Val(day) = 12 Then
    day = "B"
ElseIf Val(day) = 13 Then
    day = "C"
ElseIf Val(day) = 14 Then
    day = "E"
ElseIf Val(day) = 15 Then
    day = "F"
ElseIf Val(day) = 16 Then
    day = "G"
ElseIf Val(day) = 17 Then
    day = "H"
ElseIf Val(day) = 18 Then
    day = "J"
ElseIf Val(day) = 19 Then
    day = "K"
ElseIf Val(day) = 20 Then
    day = "L"
ElseIf Val(day) = 21 Then
    day = "M"
ElseIf Val(day) = 22 Then
    day = "N"
ElseIf Val(day) = 23 Then
    day = "P"
ElseIf Val(day) = 24 Then
    day = "R"
ElseIf Val(day) = 25 Then
    day = "S"
ElseIf Val(day) = 26 Then
    day = "T"
ElseIf Val(day) = 27 Then
    day = "V"
ElseIf Val(day) = 28 Then
    day = "W"
ElseIf Val(day) = 29 Then
    day = "X"
ElseIf Val(day) = 30 Then
    day = "Y"
ElseIf Val(day) = 31 Then
    day = "Z"
End If

If Val(month) < 10 Then
    month = Val(month)
ElseIf Val(month) = 10 Then
   month = "X"
ElseIf Val(month) = 11 Then
month = "Y"
ElseIf Val(month) = 12 Then
month = "Z"
End If
GetDateCode = day & month & year
End Function

Private Sub ShapeColorfunction(data As Integer, reg1 As Integer, reg2 As Integer, ctrl As Object)
    If (data And reg1) Then
        If (data And reg2) Then
           ctrl.BackColor = vbRed
        Else
           ctrl.BackColor = vbYellow
         End If
    ElseIf (data And reg2) Then
          ctrl.BackColor = vbGreen
    Else
          ctrl.BackColor = vbWhite
    End If
End Sub
Private Sub ShapeColorfunction1(data As Integer, reg1 As Integer, reg2 As Integer, ctrl As Shape)
    If (data And reg1) Then
        If (data And reg2) Then
           ctrl.BackColor = vbRed
        Else
           ctrl.BackColor = vbGreen
         End If
    ElseIf (data And reg2) Then
          ctrl.BackColor = vbRed
    Else
       ctrl.BackColor = vbWhite
    End If
End Sub

Private Sub ShapeColorsinglefunction(data As Integer, reg1 As Integer, ctrl As Object)
    If (data And reg1) <> 0 Then
          ctrl.BackColor = vbGreen
    Else
          ctrl.BackColor = vbYellow
    End If
End Sub
Private Sub ShapeColorsingleifunction(data As Integer, reg1 As Integer, ctrl As Object)
    If (data And reg1) <> 0 Then
          ctrl.BackColor = vbGreen
    Else
          ctrl.BackColor = vbWhite
    End If
End Sub

Private Sub Timer10_Timer()
Timer10.Enabled = False
txtScanBarcode.SetFocus
End Sub

Private Sub Timer2_Timer()
'On Error Resume Next

'    txttime = Format(Time(), "Hh:Mm:Ss")

    Static TOGGLE As Boolean
    TOGGLE = Not (TOGGLE)
    Timer2.Interval = 400
    
    With txtCommandLine
        .BorderStyle = 1
        .Alignment = 2
        .FontBold = True
       
        .FontSize = 16
    End With
       


    If Winsock1.State = 7 Then
        ShapePLCState.BackColor = vbGreen
    Else
        ShapePLCState.BackColor = vbRed
    End If
    Dim Description As String
    
    Select Case Winsock1.State
        Case 0
            Description = "Connection Closed"
        Case 1
            Description = "Connection Open"
        Case 2
            Description = "Listening For Incomming Connections"
        Case 3
            Description = "Connection Pending"
        Case 4
            Description = "Resolving Remote Host Name"
        Case 5
            Description = "Remote Host Name Successfully Resolved"
        Case 6
            Description = "Connecting-Remote Host"
        Case 7
            Description = "Connected-Remote Host"
            RetryCount = 5
        Case 8
            Description = "Connection is Closing"
        Case 9
            Description = "Connection Error"
        Case Else
            Description = "Connection Status Error"
    End Select

    
    
    If PLC_Communication_Error = True Then
       txtCommandLine.ForeColor = vbRed
       txtCommandLine.Text = "communication error"
        Exit Sub
    End If
    
    If TOGGLE = True Then
        If MsgCode >= 1 And MsgCode <= MsgCount Then
            txtCommandLine.Text = MsgText(MsgCode)

            Select Case MsgColor(MsgCode)
                Case 1
                    txtCommandLine.ForeColor = vbBlue
                Case 2
                    txtCommandLine.ForeColor = vbRed
                Case Else
                    txtCommandLine.ForeColor = vbBlack
            End Select
        Else
            txtCommandLine.Text = ""
        End If
    Else
        txtCommandLine.Text = ""
    End If

End Sub
Public Function sendEmail()
'On Error GoTo Error
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Sql As String


Sql = "Select * from Common_Set where SetType ='CommonSet'"
Set rs1 = New ADODB.Recordset
rs1.Open Sql, Con, adOpenDynamic, adLockOptimistic
If rs1("SenderEmail") <> "" And rs1("ToEmail1") <> "" And rs1("EmailBypass") = 0 Then
    Sql = "select Top 1 * from model_report_counter where MailSent = false and (DateTime < #" & Format(Now, "mm-dd-yyyy") & "# or shifttime <> '" & getShift & "')order by id desc"
    Set rs2 = New ADODB.Recordset
    rs2.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Do While rs2.EOF = False
        Dim Body As String
        Dim Subject As String
        Subject = "Production Report of Switch testing of " & rs2("ModelName") & "for date " & Format(rs2("DateTime"), "dd-mm-yyyy") & "and Shift " & rs2("ShiftTime")
        Body = "Dear Team," & vbNewLine
        Body = Body & "Below is the Production detail of Date '" & Format(rs2("DateTime"), "dd-mm-yyyy")
        Body = Body & "' and Shift '" & rs2("ShiftTime") & "' :" & vbNewLine
        Body = Body & "Model Name :- '" & rs2("ModelName") & "'" & vbNewLine
        Body = Body & "Total Ok Parts :- " & rs2("OKCounter") & vbNewLine
        Body = Body & "Total NG Parts :- " & rs2("NGCounter") & vbNewLine
        Body = Body & "Total Production Parts :- " & rs2("ProductionCounter") & vbNewLine
        If callSendEmailApi(rs1, Subject, Body) = True Then
         rs2("MailSent") = 1
         rs2.Update
        End If
        
        rs2.MoveNext
    Loop
End If
'End Function
'Error:
'ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
End Function
Private Function callSendEmailApi(rsGeneralset As ADODB.Recordset, Subject As String, Body As String) As Boolean
Dim ToEmail As String
    ToEmail = "&ToMailAddress%5b0%5d=" & rsGeneralset("ToEmail1")
    j = 0
    For i = 1 To 6
        If rsGeneralset("EmailBypass" & i) = False Then
            j = j + 1
            ReDim ToEmail1(j) As String
            ToEmail = ToEmail & "&ToMailAddress%5b" & j & "%5d=" & rsGeneralset("ToEmail" & i + 1)
        End If
    Next
    Dim URL As String
    Dim response As String
    
    URL = "http://" & rsGeneralset("WebApiLink") & "/SendMail?"
    URL = URL & "FromMailAddress=" & rsGeneralset("SenderEmail")
    URL = URL & "&FromMailPassword=" & rsGeneralset("SenderPassword")
    URL = URL & ToEmail
    URL = URL & "&subject=" & Subject
    URL = URL & "&body=" & Body
    
    Dim res As WinHttp.WinHttpRequest
    Set res = New WinHttp.WinHttpRequest
    With res
    
      ErrorLog 100, "API Initialise With URL - " & URL, "", "callsendEmailApi", ""
     .Open "Get", URL, False
     .Send
     .WaitForResponse
     response = .ResponseText
     ErrorLog 100, "API Response Recieved - " & response, "", "callsendEmailApi", ""
     If Trim(response) = "SENT" Then
     callSendEmailApi = True
     Else
     callSendEmailApi = False
     
     End If
     
    
    End With
End Function
Private Sub Load_Message_File()
On Error Resume Next
Dim iFile As Integer
Dim s As String
Dim sTextLines() As String
Dim strArray() As String
Dim WorkFile As String

    WorkFile = App.Path & "\Messages.csv"

    'Read the entire file
   iFile = FreeFile
   Open WorkFile For Input As #iFile
        s = Input(LOF(iFile), iFile)
   Close iFile
   'Split the results into separate lines
   sTextLines = Split(s, vbCrLf)

    MsgCount = UBound(sTextLines)
    ReDim MsgText(UBound(sTextLines))
    ReDim MsgColor(UBound(sTextLines))

    For i = 0 To MsgCount
        strArray = Split(sTextLines(i), ",")
        MsgText(i) = strArray(1)
        MsgColor(i) = strArray(2)
    Next

ErrorHandler:
Close iFile
End Sub
Private Sub LoadGrid()
Dim Sql As String
'DataGrid1.Col = 1
'DataGrid1.ColumnHeaders = False
'DataGrid1.RowHeight = 500

'Sql = "Select Barcode as ScanBarcode from Model_Report where ModelName='" & ModelName & "'"
'Set rsdb = New ADODB.Recordset
'rsdb.CursorLocation = adUseClient ' Set cursor location to make the recordset bookmarkable
'rsdb.Open Sql, Con, adOpenStatic, adLockOptimistic
'SaveReport
'Set DataGrid1.DataSource = rsdb
End Sub


Private Sub LoadData()

On Error GoTo Error
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim strByPass(14) As Integer
Dim j As Integer

    Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    PLcdata(2100) = Rs("ModelNo")
    PLcdata(2105) = &H1 * Val(Rs("Bypass1"))

Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
Resume Next
End Sub
Private Sub LoadSettingsData()
On Error GoTo Error
Dim Rs As ADODB.Recordset
Dim Sql As String

   Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   txtModelDesc.Text = Rs("ModelDesc")
    PartNo = Rs("PartNo")
    VendorCode = Rs("VendorCode")
    RevisionNo = Rs("RevisionNo")
   If Rs("Bypass1") = 1 Then
    ScanBypass = True
   Else
    ScanBypass = False
   End If
Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadSettingsData"
Resume Next
End Sub
Private Function getresult(pic As PictureBox) As Integer
   If pic.BackColor = vbGreen Then
    getresult = 1
   ElseIf pic.BackColor = vbRed Then
    getresult = 2
   ElseIf pic.BackColor = vbWhite Then
    getresult = 0
   End If
End Function
Private Sub SaveReport(result As String, Barcodestr As String)
'On Error GoTo Error
Dim Sql As String
Dim Rs As ADODB.Recordset
   Sql = "Select * from Model_Report"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   Rs.AddNew
      Rs("ModelName") = ModelName
      Rs("OperatorName") = LoginUser
      Rs("Date") = Format(Now(), "dd/MM/yyyy")
      Rs("Time") = Format(Now(), "hh:mm:ss")
      Rs("Barcode") = Barcodestr
      Rs("Result") = result
    Rs.Update
End Sub
Private Sub SaveToDummy()
Dim Sql As String
Dim Rs As ADODB.Recordset
   Sql = "Select * from Model_Report"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   Rs.AddNew
      Rs("ModelName") = ModelName
      Rs("OperatorName") = LoginUser
      Rs("Date") = Format(Now(), "dd/MM/yyyy")
      Rs("Time") = Format(Now(), "hh:mm:ss")
      Rs("Barcode") = barcode
      Rs("Result") = "OK"
    Rs.Update
End Sub


Private Sub SaveCounterValue()
 Dim ProdDay As String
 SaveSetting App.Title, ModelName, "OkCounter", Val(txtOKCounter.Text)
 SaveSetting App.Title, ModelName, "NGCounter", Val(txtNGCounter.Text)
 SaveSetting App.Title, ModelName, "ProductionCounter", Val(txtProductionCounter.Text)
End Sub
Private Sub SaveProductioncounter()
Dim Rs As ADODB.Recordset
Dim Sql As String
    Sql = "Select * from Model_Set where ModelName ='" & ModelName & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Rs("productioncounter") = Val(txtProductionCounter.Text)
    Rs.Update
    'txtSaveCoupler.Text = Rs("CouplerCounter")
End Sub
Private Sub GetCounterValue()
On Error Resume Next
Dim ProdDay As String
Dim Today As String
   txtOKCounter.Text = Val(GetSetting(App.Title, ModelName, "OkCounter", 0))
   txtNGCounter.Text = Val(GetSetting(App.Title, ModelName, "NgCounter", 0))
   txtProductionCounter.Text = GetSetting(App.Title, ModelName, "ProductionCounter", 0)
   runningreportdate = Format(Now(), "ddMMyy")
   runningreportshift = getShift
   tempshift = GetSetting(App.Title, ModelName, "saveshift", 0)
   tempdate = GetSetting(App.Title, ModelName, "savedate", 0)
   If tempdate <> runningreportdate Then
      If temshift <> 3 Then
      txtOKCounter.Text = 0
      txtNGCounter.Text = 0
      End If
      txtProductionCounter.Text = 0
      SaveSetting App.Title, ModelName, "saveshift", runningreportshift
      SaveSetting App.Title, ModelName, "savedate", runningreportdate
   ElseIf tempshift <> runningreportshift Then
      txtOKCounter.Text = 0
      txtNGCounter.Text = 0
      SaveSetting App.Title, ModelName, "saveshift", runningreportshift
      SaveSetting App.Title, ModelName, "savedate", runningreportdate
   End If
   SaveCounterValue
End Sub

Private Function cmdCon()
   Winsock1.Close
   Winsock1.RemoteHost = IP_HOST
   Winsock1.RemotePort = IP_PORT
   Winsock1.Connect
End Function

Private Function WinsockStstus(ByVal Value As Integer)
Dim Description As String
   Select Case Value
      Case 0
        Description = "Connection Closed"
      Case 1
        Description = "Connection Open"
      Case 2
        Description = "Listening For Incomming Connections"
      Case 3
        Description = "Connection Pending"
      Case 4
        Description = "Resolving Remote Host Name"
      Case 5
        Description = "Remote Host Name Successfully Resolved"
      Case 6
        Description = "Connecting To Remote Host"
      Case 7
        Description = "Connected To Remote Host"
        RetryCount = 0
      Case 8
        Description = "Connection is Closing"
      Case 9
        Description = "Connection Error"
      Case Else
        Description = "Connection Status Error"
   End Select
   WinsockStstus = Description
End Function

Private Sub Timer1_Timer()
   If (Winsock1.State = 7) And (CommandOn = False) Then
      Timer1.Enabled = False
      Select Case CommandType
         Case 1
            Call GetReadArray(StdReadStartAddress, StdReadCount, ReadArray)
            Winsock1.SendData ReadArray
            CVRead = CVRead + 1
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case 2
            Call GetWriteArray(StdWriteStartAddress, StdWriteCount, WriteArray)
            Winsock1.SendData WriteArray
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case 3
            Call GetReadArray((ExtendedReadStartAddress + (ExtendedReadCount * CVExtPktNo)), ExtendedReadCount, ReadArray)
            Winsock1.SendData ReadArray
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case Else
            CommandType = 1
      End Select
      Exit Sub
   Else
      Timer1.Enabled = True
      Timer1.Interval = 100
   End If

   If (Winsock1.State <> 7) Then 'And (WinSock1.State <> 6) Then
      Timer1.Interval = 1000
      Call cmdCon
   Else
      CommandOn = False
      Timer1.Interval = 1000
   End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
   LoadData
   Timer3.Interval = 150
End Sub

Private Sub Timer5_Timer()
PLC_Communication_Error = True
CommandOn = False
CommandType = 1
Timer1.Enabled = True
Timer1.Interval = 80
Timer5.Interval = 500
Timer5.Enabled = True
End Sub

Private Sub txtScanBarcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtScanBarcode.Locked = True
  txtScanBarcode.Text = Replace(txtScanBarcode.Text, "VBCR", "")
  txtScanBarcode.Text = Replace(txtScanBarcode.Text, "VBCRLF", "")
  txtScanBarcode.Text = Replace(txtScanBarcode.Text, "VBLF", "")
  If ScanValidationBypass = False Then
    If txtScanBarcode.Text = barcode Then
        PLcdata(2102) = 1
        txtScanBarcode.BackColor = vbGreen
    Else
        txtScanBarcode.BackColor = vbRed
    End If
  Else
     PLcdata(2102) = 1
     txtScanBarcode.BackColor = vbGreen
  End If
End If
End Sub
Private Function CheckBarcodeValidation(barcodetocheck As String) As Boolean
    CheckBarcodeValidation = True
    If barcodetocheck = "" Then
        PLcdata(2102) = 2
        CheckBarcodeValidation = False
    ElseIf checkpartcode(barcodetocheck) = False Then
        PLcdata(2102) = 3
        CheckBarcodeValidation = False
    ElseIf ValidateBarcodeInReport(barcodetocheck) = False Then
        PLcdata(2102) = 4
        CheckBarcodeValidation = False
    End If
End Function
Private Function checkpartcode(barcodetocheck) As Boolean
checkpartcode = True

PartNo = Mid(barcodetocheck, 1, Len(PartCode))
If PartCode <> "" Then
    If PartNo <> PartCode Then
        checkpartcode = False
    End If
End If
    
End Function

Private Function ValidateBarcodeInReport(barcodetocheck As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
   Sql = "Select * from Model_Report where barcode='" & barcodetocheck & "' and result = 1"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If Rs.EOF = False Then
      ValidateBarcodeInReport = False
   Else
      ValidateBarcodeInReport = True
   End If
End Function
Private Function ValidateBarcodeInDummy(barcodetocheck As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
   Sql = "Select * from DummyTable where barcode='" & barcodetocheck & "'"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If Rs.EOF = False Then
      ValidateBarcodeInDummy = False
   Else
      ValidateBarcodeInDummy = True
   End If
End Function
Private Function ValidateBarcodeInDummySameModel(barcodetocheck As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
   Sql = "Select * from DummyTable where barcode='" & barcodetocheck & "' and ModelName ='" & ModelName & "'"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If Rs.EOF = False Then
      ValidateBarcodeInDummySameModel = False
   Else
      ValidateBarcodeInDummySameModel = True
   End If
End Function
Private Function loaddummydata()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Row As Integer
Row = 1
   Sql = "Select * from DummyTable where ModelName ='" & ModelName & "' order by id desc"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If Rs.EOF = False Then
        txtPartScanned.Text = Rs.RecordCount
        Do While Rs.EOF = False
            VSFModel.Rows = Row + 1
            VSFModel.TextMatrix(Row, 0) = Val(txtPartScanned.Text - Row + 1)
            VSFModel.TextMatrix(Row, 1) = Rs("Barcode")
        Loop
   Else
    txtPartScanned.Text = 0
   
   End If
End Function
Private Sub deletedummyData()
On Error GoTo Error
Dim Rs As ADODB.Recordset
Dim Sql As String
   Sql = "Delete from DummyTable where ModelName ='" & ModelName & "'"
   Con.Execute Sql

Exit Sub
Error:
  ErrorLog Err.Number, Err.Description, Err.Source, "", ""
  Err.Clear

End Sub
Private Function MoveDummydataToModelReport(TreyStatus As Integer)
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim rs1 As ADODB.Recordset
Dim Sql1 As String
   
   Sql = "Select * from DummyTable where ModelName ='" & ModelName & "' order by id asc"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   Sql1 = "Select * from Model_Report"
   Set rs1 = New ADODB.Recordset
   rs1.Open Sql1, Con, adOpenDynamic, adLockOptimistic
   
        Do While Rs.EOF = False
            rs1.AddNew
            rs1("Date") = Rs("Date")
            rs1("Time") = Rs("Time")
            rs1("Barcode") = Rs("Barcode")
            rs1("Status") = Rs("Status")
            rs1("TreyStatus") = Rs("TreyStatus")
            rs1("OperatorName") = Rs("OperatorName")
            rs1("ModelName") = Rs("ModelName")
            rs1.Update
        Loop

End Function


Private Function CreateCsv()
On Error GoTo Error
Dim markingFileData As String
Dim FSO As New FileSystemObject
Dim iFile As String
Dim iFileNo As Integer

loaddummydata

For i = 1 To VSFModel.Rows - 1
    markingFileData = VSFModel.TextMatrix(i, 1) & vbCrLf
Next
If FSO.FolderExists(NotePadPath) = True Then
    iFile = NotePadPath & "\PrintFile.csv"
    If FSO.FileExists(iFile) = True Then
        FSO.DeleteFile iFile, True
    End If
    FSO.CreateTextFile iFile
    iFileNo = FreeFile
    Open iFile For Append As iFileNo
    Print #iFileNo, markingFileData
    Close iFileNo
Else
    ErrorLog 0, "Notepad FolderPath not found", "", "", ""
End If
Exit Function
Error:
  ErrorLog Err.Number, Err.Description, Err.Source, "", ""
  Err.Clear
 
End Function

Private Sub Timer9_Timer()
'rsdb.Requery
'Set DataGrid1.DataSource = rsdb
'DataGrid1.ReBind
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim SocketData() As Byte
Dim RegData, A, B, c As String
Dim i, j, K, l, M, n, ExpectedArraySize, ExtndedReadFrom, ExpectedLength As Integer
Dim Idata As Long
Dim Idata1 As Long

   Timer5.Enabled = False
   PLC_Communication_Error = False
   Winsock1.GetData SocketData
   CommandOn = False
   PlcComm = False
   Select Case CommandType
      Case 1
         K = StdReadCount * 2
         ExpectedArraySize = K + 10
         If UBound(SocketData) = ExpectedArraySize Then
            If (SocketData(0) = &HD0) And (SocketData(3) = &HFF) And (SocketData(4) = &HFF) And (SocketData(5) = 3) Then
               j = 11
               For i = StdReadStartAddress To (StdReadStartAddress + StdReadCount - 1)
                  M = CInt(SocketData(j + 1))
                  n = CInt(SocketData(j))
                  Idata = (M * 256) + n
                  If Idata > 32767 Then
                     Idata1 = Idata - 65536
                  Else
                     Idata1 = Idata
                  End If
                  PLcdata(i) = CInt(Idata1)
                  j = j + 2
               Next
               If CVRead = 1 Then CommandType = 2
               If ((CVRead >= WriteDelayCount) And ((PLcdata(StdReadStartAddress + StdReadCount - 1) = 0) Or (ExtendedRequired = False))) Then CVRead = 0
               If ((ExtendedRequired = True) And (PLcdata(StdReadStartAddress + StdReadCount - 1) > 0)) Then
                  CommandType = 3
                  CVExtPktNo = 0
               End If
               AssignPLCdata
            Else
               RejCnt = RejCnt + 1
            End If
         Else
            RejCnt = RejCnt + 1
         End If
      Case 2
         If (UBound(SocketData) = 10 And (SocketData(0) = &HD0) And (SocketData(3) = &HFF) And (SocketData(4) = &HFF) And (SocketData(5) = 3)) Then
            CommandType = 1
         Else
            RejCnt = RejCnt + 1
         End If
      Case 3
         K = ExtendedReadCount * 2
         ExpectedArraySize = K + 10
         If UBound(SocketData) = ExpectedArraySize Then
         If (SocketData(0) = &HD0) And (SocketData(3) = &HFF) And (SocketData(4) = &HFF) And (SocketData(5) = 3) Then
            j = 11
            ExtendReadFrom = ExtendedReadStartAddress + (ExtendedReadCount * CVExtPktNo)
            For i = ExtendReadFrom To (ExtendReadFrom + ExtendedReadCount - 1)
               M = CInt(SocketData(j + 1))
               n = CInt(SocketData(j))
               Idata = (M * 256) + n
               If Idata > 32767 Then
                  Idata1 = Idata - 65536
               Else
                  Idata1 = Idata
               End If
               PLcdata(i) = CInt(Idata1)
               j = j + 2
            Next
            CVExtPktNo = CVExtPktNo + 1
            If (CVExtPktNo >= NoOfExtendedPackets) Then
               CVExtPktNo = 0
               If (CVRead = 1) Then
                  CommandType = 2
               Else
                  CommandType = 1
               End If
               If ((CVRead >= WriteDelayCount)) Then CVRead = 0
            End If
         Else
            RejCnt = RejCnt + 1
         End If
      Else
         RejCnt = RejCnt + 1
      End If
   End Select
 
   ' txtModelName = CommandType
   ' txtOd4 = UBound(SocketData)
   'Text2 = CommandType & "+" & CVExtPktNo
   Timer1.Interval = 10
   Timer1.Enabled = True
End Sub
