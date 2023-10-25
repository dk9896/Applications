VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMonitor 
   Caption         =   "22_02"
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
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   Begin VB.Timer Timer14 
      Left            =   9480
      Top             =   5280
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
      Height          =   10575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   19815
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         Height          =   4695
         Left            =   7560
         ScaleHeight     =   4635
         ScaleWidth      =   8115
         TabIndex        =   46
         Top             =   6000
         Visible         =   0   'False
         Width           =   8175
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
            Picture         =   "frmMonitor.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   3480
            UseMaskColor    =   -1  'True
            Width           =   1275
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
            TabIndex        =   62
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
            TabIndex        =   61
            Text            =   "0000"
            Top             =   2880
            Width           =   855
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
            TabIndex        =   60
            Text            =   "0000"
            Top             =   2880
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
            TabIndex        =   59
            Text            =   "0000"
            Top             =   2880
            Width           =   975
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
            TabIndex        =   58
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
            TabIndex        =   57
            Text            =   "0000"
            Top             =   2160
            Width           =   855
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
            TabIndex        =   56
            Text            =   "0000"
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
            Index           =   2
            Left            =   1800
            TabIndex        =   55
            Text            =   "0000"
            Top             =   2160
            Width           =   975
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
            TabIndex        =   54
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
            TabIndex        =   53
            Text            =   "0000"
            Top             =   1440
            Width           =   855
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
            TabIndex        =   52
            Text            =   "0000"
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
            Index           =   1
            Left            =   1800
            TabIndex        =   51
            Text            =   "0000"
            Top             =   1440
            Width           =   975
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
            TabIndex        =   50
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
            TabIndex        =   49
            Text            =   "0000"
            Top             =   720
            Width           =   855
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
            TabIndex        =   48
            Text            =   "0000"
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
            Index           =   0
            Left            =   1800
            TabIndex        =   47
            Text            =   "0000"
            Top             =   720
            Width           =   975
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
            TabIndex        =   70
            Top             =   840
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
            TabIndex        =   69
            Top             =   1560
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
            TabIndex        =   68
            Top             =   2280
            Width           =   1335
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
            TabIndex        =   67
            Top             =   3000
            Width           =   1335
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
            TabIndex        =   65
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
            TabIndex        =   64
            Top             =   240
            Width           =   735
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
            TabIndex        =   63
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox PictureBreakdown 
         BackColor       =   &H80000010&
         Height          =   6015
         Left            =   -1200
         ScaleHeight     =   5955
         ScaleWidth      =   8595
         TabIndex        =   28
         Top             =   8040
         Visible         =   0   'False
         Width           =   8655
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
            Picture         =   "frmMonitor.frx":0C42
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   4680
            UseMaskColor    =   -1  'True
            Width           =   1275
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
            TabIndex        =   32
            Top             =   4440
            Width           =   4575
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
            TabIndex        =   31
            Top             =   2760
            Width           =   1935
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
            TabIndex        =   29
            Top             =   840
            Width           =   1815
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
            TabIndex        =   34
            Top             =   4800
            Width           =   2295
         End
      End
      Begin VB.TextBox txtOKCounter 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   0
         Left            =   18600
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3360
         Width           =   990
      End
      Begin VB.TextBox txtNGCounter 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   0
         Left            =   18600
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2760
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         TabIndex        =   40
         Top             =   8400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5760
         TabIndex        =   39
         Text            =   "Text3"
         Top             =   480
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   615
         Left            =   4320
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtproductioncounter 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0"
         Top             =   9120
         Width           =   2490
      End
      Begin VB.TextBox txtTargetProduction 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9480
         TabIndex        =   36
         Top             =   8400
         Width           =   1215
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
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   2235
         TabIndex        =   26
         Top             =   240
         Width           =   2295
         Begin VB.CheckBox chkLoadCellCalibration 
            Caption         =   "Load Cell Calibration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   45
            Top             =   0
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   735
            Left            =   0
            Picture         =   "frmMonitor.frx":1884
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.TextBox txtBarcode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   7560
         Width           =   6135
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
         Left            =   18360
         TabIndex        =   15
         Top             =   9600
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
            Picture         =   "frmMonitor.frx":440A
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   1275
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
         Height          =   1815
         Left            =   17400
         TabIndex        =   12
         Top             =   7680
         Visible         =   0   'False
         Width           =   2295
         Begin VB.Label lblGo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   72
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1425
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.Label lblNg 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   72
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1665
            Left            =   60
            TabIndex        =   13
            Top             =   120
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.Frame Frame8 
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
         Height          =   615
         Left            =   17160
         TabIndex        =   9
         Top             =   1440
         Width           =   2535
         Begin VB.TextBox txtBatchCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Batch Count"
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
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
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
         Height          =   1215
         Left            =   17160
         TabIndex        =   6
         Top             =   240
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
            TabIndex        =   18
            Top             =   360
            Width           =   720
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
            TabIndex        =   20
            Top             =   480
            Width           =   375
         End
         Begin VB.Shape shapeInternet 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   1320
            Top             =   840
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
            TabIndex        =   17
            Top             =   360
            Width           =   1575
         End
         Begin VB.Shape ShapePLCState 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   1320
            Top             =   0
            Width           =   855
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
            TabIndex        =   8
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Internet Con"
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
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1050
         End
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
         Height          =   510
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmMonitor.frx":504C
         Top             =   9840
         Width           =   18015
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
         Left            =   12240
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.Timer Timer8 
            Interval        =   60000
            Left            =   840
            Top             =   240
         End
         Begin VB.Timer Timer7 
            Left            =   360
            Top             =   1320
         End
         Begin VB.Timer Timer13 
            Enabled         =   0   'False
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer12 
            Enabled         =   0   'False
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer11 
            Enabled         =   0   'False
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer6 
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer3 
            Left            =   840
            Top             =   960
         End
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   120
            Top             =   960
         End
         Begin VB.Timer Timer2 
            Left            =   480
            Top             =   960
         End
         Begin VB.Timer Timer4 
            Left            =   1320
            Top             =   960
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
            TabIndex        =   3
            Top             =   360
            Width           =   1440
         End
         Begin VB.Timer Timer5 
            Left            =   2640
            Top             =   1080
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
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "MODEL DESC"
         Top             =   240
         Width           =   14655
      End
      Begin VB.Frame frmCoupler 
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
         Height          =   735
         Index           =   0
         Left            =   17160
         TabIndex        =   71
         Top             =   1920
         Width           =   2535
         Begin VB.TextBox txtCouplerCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   0
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   72
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Coupler Count"
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
            Index           =   0
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox txtposition 
         Alignment       =   2  'Center
         Height          =   555
         Index           =   1
         Left            =   9480
         TabIndex        =   84
         Text            =   "000.0"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtcurrent 
         Alignment       =   2  'Center
         Height          =   555
         Index           =   1
         Left            =   9480
         TabIndex        =   86
         Text            =   "0.000"
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox txtmvd 
         Alignment       =   2  'Center
         Height          =   555
         Index           =   1
         Left            =   9480
         TabIndex        =   88
         Text            =   "000"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox txtTestVolt 
         Alignment       =   2  'Center
         Height          =   555
         Left            =   6480
         TabIndex        =   76
         Text            =   "00.00"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtposition 
         Alignment       =   2  'Center
         Height          =   555
         Index           =   0
         Left            =   3240
         TabIndex        =   78
         Text            =   "000.0"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtcurrent 
         Alignment       =   2  'Center
         Height          =   555
         Index           =   0
         Left            =   3240
         TabIndex        =   80
         Text            =   "0.000"
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox txtmvd 
         Alignment       =   2  'Center
         Height          =   555
         Index           =   0
         Left            =   3240
         TabIndex        =   82
         Text            =   "000"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "MVD"
         Height          =   495
         Left            =   7680
         TabIndex        =   87
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Current"
         Height          =   495
         Left            =   7680
         TabIndex        =   85
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Position"
         Height          =   495
         Left            =   7680
         TabIndex        =   83
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "MVD"
         Height          =   495
         Left            =   840
         TabIndex        =   81
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Current"
         Height          =   495
         Left            =   840
         TabIndex        =   79
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Position"
         Height          =   495
         Left            =   840
         TabIndex        =   77
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Test Volt"
         Height          =   375
         Left            =   4200
         TabIndex        =   75
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label11 
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
         Left            =   17280
         TabIndex        =   42
         Top             =   3480
         Width           =   1215
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
         Left            =   17280
         TabIndex        =   41
         Top             =   2880
         Width           =   855
      End
      Begin VB.Image ImgPart 
         Height          =   3975
         Left            =   14760
         Stretch         =   -1  'True
         Top             =   5520
         Width           =   4935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   43
         Left            =   7680
         TabIndex        =   25
         Top             =   9240
         Width           =   1665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Production"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   40
         Left            =   7680
         TabIndex        =   24
         Top             =   8520
         Width           =   1530
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   7680
         TabIndex        =   21
         Top             =   7680
         Width           =   720
      End
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   -1800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   495
      Left            =   9480
      TabIndex        =   66
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   9480
      TabIndex        =   33
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ILLumination Curr. LH "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   34
      Left            =   7920
      TabIndex        =   23
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   7920
      TabIndex        =   22
      Top             =   7320
      Width           =   375
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

'----------------
Dim PLC_Communication_Error As Boolean
Dim MsgText() As String
Dim MsgColor() As Integer
Dim MsgCount As Integer
Dim CloseScreen As Boolean
Dim runningreportdate As Date
Dim runningreportshift As String
Dim ModelNo As Integer
Private Declare Function InternetGetConnectedState Lib _
    "wininet" (ByRef dwflags As Long, ByVal dwReserved As _
    Long) As Long
Dim pulseCalibration(4) As Boolean
  
Private Sub chkLoadCellCalibration_Click()
If chkLoadCellCalibration.Value = 1 Then
Picture2.Visible = True
Else
Picture2.Visible = False
End If
End Sub

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

Private Sub cmdclosebreakdownscreen_Click()
    PictureBreakdown.Visible = False
    Command2.Enabled = True
End Sub

Private Sub cmdCloseLoadCalibration_Click()
chkLoadCellCalibration.Value = False
End Sub

Private Sub cmdfullbreakdown_Click()
    cmdrunningbreakdown.Enabled = False
    cmdfullbreakdown.Enabled = False
    cmdgolive.Enabled = True
    'cmdclosebreakdownscreen.Enabled = False
    SaveBreakDown 3, 1
    PLcdata(348) = 3
End Sub

Private Sub cmdgolive_Click()
    cmdrunningbreakdown.Enabled = True
    cmdfullbreakdown.Enabled = True
    cmdgolive.Enabled = False
    'cmdclosebreakdownscreen.Enabled = True
    SaveBreakDown 1, 0
    PLcdata(348) = 1
End Sub

Private Sub cmdrunningbreakdown_Click()
    cmdrunningbreakdown.Enabled = False
    cmdfullbreakdown.Enabled = False
    cmdgolive.Enabled = True
    'cmdclosebreakdownscreen.Enabled = False
    SaveBreakDown 2, 1
    PLcdata(348) = 2
End Sub

Private Sub cmdSetCalibration_Click(Index As Integer)
If Index = 0 Then
    pulseCalibration(0) = True
    If Val(txtCalibrateValue(0).Text) > -32000 And Val(txtCalibrateValue(0).Text) < 32000 Then
    PLcdata(411) = Val(txtCalibrateValue(0).Text)
    End If
    Timer11.Interval = 1000
    Timer11.Enabled = True
ElseIf Index = 1 Then
    pulseCalibration(1) = True
    If Val(txtCalibrateValue(1).Text) > -32000 And Val(txtCalibrateValue(1).Text) < 32000 Then
    PLcdata(412) = Val(txtCalibrateValue(1).Text)
    End If
    Timer12.Interval = 1000
    Timer12.Enabled = True
ElseIf Index = 2 Then
    pulseCalibration(2) = True
    If Val(txtCalibrateValue(2).Text) > -32000 And Val(txtCalibrateValue(2).Text) < 32000 Then
    PLcdata(413) = Val(txtCalibrateValue(2).Text)
    End If
    Timer13.Interval = 1000
    Timer13.Enabled = True
ElseIf Index = 3 Then
    pulseCalibration(3) = True
    If Val(txtCalibrateValue(3).Text) > -32000 And Val(txtCalibrateValue(3).Text) < 32000 Then
    PLcdata(414) = Val(txtCalibrateValue(3).Text)
    End If
    Timer14.Interval = 1000
    Timer14.Enabled = True
End If
End Sub

Private Sub Command1_Click()
  If Val(txtTargetProduction.Text) > 0 Then
      Command1.Visible = False
      txtTargetProduction.Enabled = False
      txtTargetProduction.BackColor = vbWhite
      runningreportshift = getShift
      runningreportdate = TempReportDate
      SaveSetting App.Title, ModelName, "TargetProduction", txtTargetProduction.Text
      GetCounterValue
      PLcdata(349) = 0
  Else
    txtTargetProduction.BackColor = vbRed
  End If
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    PictureBreakdown.Visible = True


End Sub


Private Sub Command3_Click()
'PLcdata(109) = Val(Text3.Text)
'txtproductioncounter.Text = 1
'SaveCounter
PLcdata(104) = 2
AssignPLCdata
End Sub

Private Sub Command7_Click()

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
Dim rs As ADODB.Recordset

   'To Load Com port in Monitor
   Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   Dim ComPort(3) As Integer
   Dim ComPortBP(3) As Integer
   ComPort(1) = rs("ComPort1")
''    ComPort(2) = Rs("ComPort2")
    ComPortBP(1) = rs("ComPortBP1")
''      ComPortBP(2) = Rs("ComPortBP2")
   PrinterName = rs("PrinterName1")
   Initialise
   WinSock1.Protocol = sckTCPProtocol
   'txtIP.Text = WinSock1.LocalIP
   IP_HOST = rs("PLC_IP") '"192.168.1.30"
   IP_PORT = rs("PLC_Port")
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
runningreportshift = GetSetting(App.Title, ModelName, "saveshift", 0)
runningreportdate = GetSetting(App.Title, ModelName, "savedate", 0)
PLcdata(340) = 1
GetCounterValue
ConnectToPLC
Timer1.Enabled = True
Timer1.Interval = 1000
Timer2.Enabled = True
Timer2.Interval = 1000
Timer3.Interval = 500
Timer3.Enabled = True
'txtDate.Text = Date
'txttime.Text = Format(Time(), "hh:mm:ss")
'txtOperName.Text = LoginUser
LoadGrid
Pulse = False
Exit Sub
End Sub

Private Sub UserAccess()
   If AccessType = "0" Then 'Disable or Hide For Operator
      'CmdOKCounter.Visible = False
      'CmdNgCounter.Visible = False
      'Command1.Visible = False
      chkLoadCellCalibration.Visible = False
   ElseIf AccessType = "1" Then 'Disable or Hide for AccessType 1
      'CmdOKCounter.Visible = False
      'CmdNgCounter.Visible = False
      'Command1.Visible = False
      chkLoadCellCalibration.Visible = False
   ElseIf AccessType = "2" Then 'Show All Which Will Disable or Hide For One
      'CmdOKCounter.Visible = True
      'CmdNgCounter.Visible = True
   End If
End Sub

Private Function AssignPLCdata()
On Error GoTo Error
   MsgCode = PLcdata(108)
      
   
   ShapeColorfunction PLcdata(110), &H1, &H2, txtTestVolt
   
   ShapeColorfunction PLcdata(111), &H1, &H2, txtposition(0)
   ShapeColorfunction PLcdata(111), &H4, &H8, txtcurrent(0)
   ShapeColorfunction PLcdata(111), &H10, &H20, txtmvd(0)
   
   ShapeColorfunction PLcdata(111), &H1, &H2, txtposition(1)
   ShapeColorfunction PLcdata(111), &H4, &H8, txtcurrent(1)
   ShapeColorfunction PLcdata(111), &H10, &H20, txtmvd(1)
   
   txtTestVolt.Text = Format(PLcdata(120), "0")
   
   txtposition(0).Text = Format(PLcdata(125), "0")
   txtcurrent(0).Text = Format(PLcdata(126), "0")
   txtmvd(0).Text = Format(PLcdata(127), "0")
   
   txtposition(1).Text = Format(PLcdata(130), "0")
   txtcurrent(1).Text = Format(PLcdata(131), "0")
   txtmvd(1).Text = Format(PLcdata(132), "0")
   
   txtCycleTime.Text = Format(PLcdata(107) / 10, "0.0")
   
   If PLcdata(181) = 0 And pulseBreakdown = True Then
      pulseBreakdown = False
      'PictureBreakdown.Visible = False
   ElseIf PLcdata(181) = 1 And pulseBreakdown = False Then
      pulseBreakdown = True
      PictureBreakdown.Visible = True
      cmdrunningbreakdown.Enabled = False
      cmdfullbreakdown.Enabled = False
      cmdgolive.Enabled = True
      'cmdclosebreakdownscreen.Enabled = False
      
   ElseIf PLcdata(181) = 2 And pulseBreakdown = False Then
      pulseBreakdown = True
      PictureBreakdown.Visible = True
      cmdrunningbreakdown.Enabled = False
      cmdfullbreakdown.Enabled = False
      cmdgolive.Enabled = True
      'cmdclosebreakdownscreen.Enabled = False
   End If
   
   If PLcdata(182) = 0 And PulseScan = False Then
      PulseScan = True
      txtBarcode.Locked = False
      txtBarcode.BackColor = vbWhite
      txtBarcode.Text = ""
      txtBarcode.Locked = True
      PLcdata(350) = 0
   ElseIf PLcdata(182) = 1 And PulseScan = True Then
      PulseScan = False
      txtBarcode.Locked = False
      txtBarcode.BackColor = vbYellow
      txtBarcode.SetFocus
      
   End If
   If PLcdata(100) = 0 And pulse1 = False Then
      pulse1 = True
      lblGo.Visible = False
      lblNg.Visible = False
   ElseIf PLcdata(100) = 1 And pulse1 = True Then
      pulse1 = False
      lblGo.Visible = True
      GetCounterValue
      txtproductioncounter.Text = Val(txtproductioncounter.Text) + 1
      txtOKCounter(0).Text = Val(txtOKCounter(0).Text) + 1
      txtBatchCounter.Text = Val(txtBatchCounter.Text) + 1
      txtTargetProduction.Text = Val(txtTargetProduction.Text) - 1
      txtCouplerCounter(0).Text = Val(txtCouplerCounter(0).Text) + 1
      If pulsePrinterBypass = False Then
        PrintLabel JustPrinter1
      End If
      SaveProductioncounter
      SaveReport 1, 1
      SaveCounter
      SaveCounterValue
   ElseIf PLcdata(100) = 2 And pulse1 = True Then
      pulse1 = False
      GetCounterValue
      lblNg.Visible = True
      txtNGCounter(0).Text = Val(txtNGCounter(0).Text) + 1
      SaveReport 2, 1
      SaveCounter
      SaveCounterValue
   End If
      
Exit Function
Error:
   ErrorLog Err.Number, Err.Description & "---", Erl, Me.Name, "Assign PLC Data"
   Resume Next
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
Private Sub ShapeColorfunction1(data As Integer, reg1 As Integer, reg2 As Integer, ctrl As VSFlexGrid, Row As Integer, Col As Integer)
    If (data And reg1) Then
        If (data And reg2) Then
           ctrl.Cell(flexcpBackColor, Row, Col) = vbRed
        Else
           ctrl.Cell(flexcpBackColor, Row, Col) = vbGreen
         End If
    ElseIf (data And reg2) Then
          ctrl.Cell(flexcpBackColor, Row, Col) = vbRed
    Else
          ctrl.Cell(flexcpBackColor, Row, Col) = vbWhite
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

Private Sub Timer11_Timer()

Timer11.Enabled = False
pulseCalibration(0) = False
End Sub

Private Sub Timer12_Timer()

Timer12.Enabled = False
pulseCalibration(1) = False
End Sub

Private Sub Timer13_Timer()

Timer13.Enabled = False
pulseCalibration(2) = False
End Sub

Private Sub Timer14_Timer()
Timer14.Enabled = False
pulseCalibration(3) = False
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
       
    If InternetGetConnectedState(0, 0) = 1 Then
        shapeInternet.BackColor = vbGreen
        'sendEmail
    Else
        shapeInternet.BackColor = vbRed
    End If
    
    'Text1.Text = WinsockStstus(WinSock1.State)


    If WinSock1.State = 7 Then
        ShapePLCState.BackColor = vbGreen
    Else
        ShapePLCState.BackColor = vbRed
    End If
    Dim Description As String
    
    Select Case WinSock1.State
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
On Error GoTo Error

With VSCavity1
  
  .Cols = 3
  .Rows = 6
  .RowHeightMin = 500
  .RowHeightMax = 1000
  .ColAlignment(0) = flexAlignCenterCenter

   .TextMatrix(0, 0) = "Process"
   .TextMatrix(0, 1) = "On to Off"
   .TextMatrix(0, 2) = "Of to On"

   .TextMatrix(1, 0) = "Testing Volt(V)"
   .TextMatrix(2, 0) = "Testing Cur.(A)"
   .TextMatrix(3, 0) = "MVD(V)"
   .TextMatrix(4, 0) = "Cont. Break Point(mm)"
   .TextMatrix(5, 0) = "Operating Force(gF)"
   

    For Col = 1 To .Cols - 1
        .FixedAlignment(Col) = flexAlignCenterCenter
        .ColAlignment(Col) = flexAlignCenterCenter
        .ColWidth(Col) = 1550
    Next
    
    .ColWidth(0) = 2500

 End With
 With VSCavity2
  
  .Cols = 3
  .Rows = 6
  .RowHeightMin = 500
  .RowHeightMax = 1000
  .ColAlignment(0) = flexAlignCenterCenter

   .TextMatrix(0, 0) = "Process"
   .TextMatrix(0, 1) = "On to Off"
   .TextMatrix(0, 2) = "Off to On"


   .TextMatrix(1, 0) = "Testing Volt(V)"
   .TextMatrix(2, 0) = "Testing Cur.(A)"
   .TextMatrix(3, 0) = "MVD(V)"
   .TextMatrix(4, 0) = "Cont. Break Point(mm)"
   .TextMatrix(5, 0) = "Operating Force(gF)"
   
    For Col = 1 To .Cols - 1
        .FixedAlignment(Col) = flexAlignCenterCenter
        .ColAlignment(Col) = flexAlignCenterCenter
        .ColWidth(Col) = 1550
    Next
    
    .ColWidth(0) = 2500

 End With
 With VSCavity3
  
  .Cols = 3
  .Rows = 6
  .RowHeightMin = 500
  .RowHeightMax = 1000
  .ColAlignment(0) = flexAlignCenterCenter

   .TextMatrix(0, 0) = "Process"
   .TextMatrix(0, 1) = "On to Off"
   .TextMatrix(0, 2) = "Off to On"

   
   .TextMatrix(1, 0) = "Testing Volt(V)"
   .TextMatrix(2, 0) = "Testing Cur.(A)"
   .TextMatrix(3, 0) = "MVD(V)"
   .TextMatrix(4, 0) = "Cont. Break Point(mm)"
   .TextMatrix(5, 0) = "Operating Force(gF)"
   
    For Col = 1 To .Cols - 1
        .FixedAlignment(Col) = flexAlignCenterCenter
        .ColAlignment(Col) = flexAlignCenterCenter
        .ColWidth(Col) = 1550
    Next
    
    .ColWidth(0) = 2500

 End With
 With VSCavity4
  
  .Cols = 3
  .Rows = 6
  .RowHeightMin = 500
  .RowHeightMax = 1000
  .ColAlignment(0) = flexAlignCenterCenter

   .TextMatrix(0, 0) = "Process"
   .TextMatrix(0, 2) = "Off to On"
   .TextMatrix(0, 1) = "On to Off"

   
   .TextMatrix(1, 0) = "Testing Volt(V)"
   .TextMatrix(2, 0) = "Testing Cur.(A)"
   .TextMatrix(3, 0) = "MVD(V)"
   .TextMatrix(4, 0) = "Cont. Break Point(mm)"
   .TextMatrix(5, 0) = "Operating Force(gF)"
   
    For Col = 1 To .Cols - 1
        .FixedAlignment(Col) = flexAlignCenterCenter
        .ColAlignment(Col) = flexAlignCenterCenter
        .ColWidth(Col) = 1550
    Next
    
    .ColWidth(0) = 2500

 End With
Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadGrid"
Resume Next
End Sub
Private Sub LoadData()

On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
Dim strByPass(14) As Integer
Dim j As Integer

    Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    'txtModelDesc.Text = Trim(Rs("ModelDesc"))
    If ((Val(txtCouplerCounter(0).Text) >= Val(rs("CouplerCounter1"))) And Val(rs("Bypass7")) = 0) Then
        PLcdata(335) = 1
        txtCouplerCounter(0).BackColor = vbRed
    ElseIf (Val(txtCouplerCounter(1).Text) >= Val(rs("CouplerCounter2"))) And Val(rs("Bypass8")) = 0 Then
        PLcdata(335) = 1
        txtCouplerCounter(1).BackColor = vbRed
    ElseIf (Val(txtCouplerCounter(2).Text) >= Val(rs("CouplerCounter3"))) And Val(rs("Bypass9")) = 0 Then
        PLcdata(335) = 1
        txtCouplerCounter(2).BackColor = vbRed
    ElseIf (Val(txtCouplerCounter(3).Text) >= Val(rs("CouplerCounter4"))) And Val(rs("Bypass10")) = 0 Then
        PLcdata(335) = 1
        txtCouplerCounter(3).BackColor = vbRed
    ElseIf Val(txtBatchCounter.Text) >= setBatchCounter Then
        PLcdata(335) = 2
    Else
        txtCouplerCounter(0).BackColor = vbWhite
        txtCouplerCounter(1).BackColor = vbWhite
        txtCouplerCounter(2).BackColor = vbWhite
        txtCouplerCounter(3).BackColor = vbWhite
        PLcdata(335) = 0
    End If
    If chkLoadCellCalibration.Value = 1 Then
        PLcdata(340) = 3
    Else
        PLcdata(340) = 1
    End If
    j = 0
    K = 1
    A = "On"
    For i = 1 To 8
      
      PLcdata(200 + j) = Val(rs("TestVol" & A & "min" & K)) * 100
      PLcdata(201 + j) = Val(rs("TestCur" & A & "min" & K)) * 100
      PLcdata(202 + j) = Val(rs("MVD" & A & "min" & K)) * 100
      PLcdata(203 + j) = Val(rs("ContBrkPnt" & A & "min" & K)) * 100
      PLcdata(204 + j) = Val(rs("OpeForce" & A & "min" & K))
'      PLcdata(205 + j) = Val(rs("ContRes" & A & "min" & K))
      PLcdata(206 + j) = Val(rs("TestVol" & A & "max" & K)) * 100
      PLcdata(207 + j) = Val(rs("TestCur" & A & "max" & K)) * 100
      PLcdata(208 + j) = Val(rs("MVD" & A & "max" & K)) * 100
      PLcdata(209 + j) = Val(rs("ContBrkPnt" & A & "max" & K)) * 100
      PLcdata(210 + j) = Val(rs("OpeForce" & A & "max" & K))
'      PLcdata(211 + j) = Val(rs("ContRes" & A & "max" & K))
      If A = "On" Then A = "Off" Else A = "On"
      j = j + 15
      K = i \ 2 + 1
    Next
    j = 355
    For i = 1 To 4
      
      PLcdata(j) = Val(rs("TestVolOnTarget" & i)) * 100
      PLcdata(j + 1) = Val(rs("TestCurOnTarget" & i)) * 100
      PLcdata(j + 2) = Val(rs("MVDOnTarget" & i)) * 100
      PLcdata(j + 3) = Val(rs("ContBrkPntOnTarget" & i)) * 100
      PLcdata(j + 4) = Val(rs("OpeForceOnTarget" & i))
'      PLcdata(j + 5) = Val(rs("ContResOnTarget" & i))
      PLcdata(j + 6) = Val(rs("TestVolOffTarget" & i)) * 100
      PLcdata(j + 7) = Val(rs("TestCurOffTarget" & i)) * 100
      PLcdata(j + 8) = Val(rs("MVDOffTarget" & i)) * 100
      PLcdata(j + 9) = Val(rs("ContBrkPntOffTarget" & i)) * 100
      PLcdata(j + 10) = Val(rs("OpeForceOffTarget" & i))
'      PLcdata(j + 11) = Val(rs("ContResOffTarget" & i))
      j = j + 12
    Next
    PLcdata(333) = Val(rs("ServoHomePos")) * 1000
    PLcdata(334) = Val(rs("ServoHomespeed")) * 1000
    PLcdata(336) = Val(rs("ServoTestPos")) * 1000
    PLcdata(337) = Val(rs("ServoTestSpeed")) * 1000
    PLcdata(338) = Val(rs("ServoFastPos")) * 1000
    PLcdata(339) = Val(rs("ServoFastSpeed")) * 1000
    ModelNo = rs("ModelNo")
    PLcdata(331) = rs("ModelNo")
    PLcdata(330) = 0
    PLcdata(330) = PLcdata(330) + &H1 * Val(rs("Bypass1"))
    PLcdata(330) = PLcdata(330) + &H2 * Val(rs("Bypass2"))
    PLcdata(330) = PLcdata(330) + &H4 * Val(rs("Bypass3"))
    PLcdata(330) = PLcdata(330) + &H8 * Val(rs("Bypass4"))
    PLcdata(330) = PLcdata(330) + &H10 * Val(rs("Bypass5"))
    PLcdata(330) = PLcdata(330) + &H20 * Val(rs("Bypass6"))
    PLcdata(330) = PLcdata(330) + &H40 * Val(rs("Bypass7"))
    PLcdata(330) = PLcdata(330) + &H80 * Val(rs("ByPass8"))
    PLcdata(330) = PLcdata(330) + &H100 * Val(rs("ByPass9"))
    PLcdata(330) = PLcdata(330) + &H200 * Val(rs("ByPass10"))
    PLcdata(330) = PLcdata(330) + &H400 * Val(rs("ByPass11"))
    PLcdata(330) = PLcdata(330) + &H800 * Val(rs("ByPass12"))
    If Val(rs("bypass7")) = 1 Then
        VSCavity1.Visible = False
        Text2.Visible = False
    End If
    If Val(rs("bypass8")) = 1 Then
        VSCavity2.Visible = False
        Text4.Visible = False
    End If
    If Val(rs("bypass9")) = 1 Then
        VSCavity3.Visible = False
        Text5.Visible = False
    End If
    If Val(rs("bypass10")) = 1 Then
        VSCavity4.Visible = False
        Text6.Visible = False
    End If
    PLcdata(410) = 0
    If pulseCalibration(0) = True Then
        PLcdata(410) = PLcdata(410) + &H1
    ElseIf pulseCalibration(1) = True Then
    
        PLcdata(410) = PLcdata(410) + &H2
    ElseIf pulseCalibration(2) = True Then
    
        PLcdata(410) = PLcdata(410) + &H4
    ElseIf pulseCalibration(3) = True Then
    
        PLcdata(410) = PLcdata(410) + &H8
    End If
    
    chkproductioncount
Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
Resume Next
End Sub
Private Sub chkproductioncount()
    tempgetshift = getShift
    'TempReportDate
       tempshift = GetSetting(App.Title, ModelName, "saveshift", 0)
       TempDate = GetSetting(App.Title, ModelName, "savedate", 0)
       If Val(txtTargetProduction.Text) > 0 And txtTargetProduction.BackColor <> vbYellow Then
        If TempReportDate <> DateValue(TempDate) Then
            txtTargetProduction.Enabled = True
            txtTargetProduction.Text = ""
            txtTargetProduction.SetFocus
            txtTargetProduction.BackColor = vbYellow
            Command1.Visible = True
            PLcdata(349) = 1
            Exit Sub
'        Else
'            If tempgetshift <> tempshift Then
'                txtTargetProduction.Locked = False
'                txtTargetProduction.Text = ""
'                txtTargetProduction.SetFocus
'                txtTargetProduction.BackColor = vbYellow
'                Command1.Visible = True
'                PLcdata(349) = 1
'                Exit Sub
'            End If
        End If
    ElseIf txtTargetProduction.BackColor <> vbYellow Then
        txtTargetProduction.Locked = False
        txtTargetProduction.Text = ""
        txtTargetProduction.SetFocus
        txtTargetProduction.BackColor = vbYellow
        Command1.Visible = True
        PLcdata(349) = 1
        
    End If
End Sub
Private Sub LoadSettingsData()
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String

   Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   txtModelDesc.Text = rs("ModelDesc")
    PartNo = rs("PrintPartNo")
    'BarcodeLength = Rs("BarcodeLength")
    HardwareNo = rs("HardwareNo")
    SerialStartingtxt = rs("SerialStartingtxt")
    setBatchCounter = rs("BatchCounter")
    'setCouplerCounter = rs("CouplerCounter1")
    VendorId = rs("VendorId")
    'PrintSwitchName = Rs("PrintSwitchName")
    'PrintLineCode = Rs("PrintLineCode")
    
    ImgPart.Picture = LoadPicture(rs("PartImage"))
    txtproductioncounter.Text = rs("productioncounter")
    If Val(rs("bypass7")) = 1 Then
        VSCavity1.Visible = False
        Text2.Visible = False
        txtOKCounter(0).Visible = False
        txtNGCounter(0).Visible = False
        Label11.Visible = False
        Label10.Visible = False
        frmCoupler(0).Visible = False
        lblContact1.Visible = False
        shpSwContact1.Visible = False
    End If
    If Val(rs("bypass8")) = 1 Then
        VSCavity2.Visible = False
        Text4.Visible = False
        txtOKCounter(1).Visible = False
        txtNGCounter(1).Visible = False
        Label6(0).Visible = False
        Label8(0).Visible = False
        frmCoupler(1).Visible = False
        lblContact2.Visible = False
        shpSwContact2.Visible = False
    End If
    If Val(rs("bypass9")) = 1 Then
        VSCavity3.Visible = False
        Text5.Visible = False
        txtOKCounter(2).Visible = False
        txtNGCounter(2).Visible = False
        Label6(1).Visible = False
        Label8(1).Visible = False
        frmCoupler(2).Visible = False
        lblContact3.Visible = False
        shpSwContact3.Visible = False
    End If
    If Val(rs("bypass10")) = 1 Then
        VSCavity4.Visible = False
        Text6.Visible = False
        txtOKCounter(3).Visible = False
        txtNGCounter(3).Visible = False
        Label6(2).Visible = False
        Label8(2).Visible = False
        frmCoupler(3).Visible = False
        lblContact4.Visible = False
        shpSwContact4.Visible = False
    End If
    
    If Val(rs("PrinterBypass")) = 1 Then
        pulsePrinterBypass = True
        txtBarcode.Visible = False
        Label4(8).Visible = False
    Else
        pulsePrinterBypass = False
    End If
    If Val(rs("Bypass12")) = 1 Then
      txtCycleTime.Visible = False
      Label5.Visible = False
      Label3.Visible = False
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
Private Sub SaveReport(result As String, Cavity As Integer)
'On Error GoTo Error
Dim Sql As String
Dim rs As ADODB.Recordset
   Sql = "Select * from Model_Report"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   rs.AddNew
      rs("ModelName") = ModelName
      rs("OperatorName") = LoginUser
      rs("Date") = Format(Now(), "dd/MM/yyyy")
      rs("Time") = Format(Now(), "hh:mm:ss")
      rs("Barcode") = barcode
      rs("Result") = result
      rs("Cavity") = Cavity
    If Cavity = 1 Then
     With VSCavity1
     Col = 2
     rs("TestVolOn") = .TextMatrix(1, Col)
     rs("TestCurOn") = .TextMatrix(2, Col)
     rs("MVDOn") = .TextMatrix(3, Col)
     rs("ContBrkPntOn") = .TextMatrix(4, Col)
     rs("OpeForceOn") = .TextMatrix(5, Col)
     'rs("ContResOn1") = .TextMatrix(6, Col)
     Col = 1
     rs("TestVolOff") = .TextMatrix(1, Col)
     rs("TestCurOff") = .TextMatrix(2, Col)
     rs("MVDOff") = .TextMatrix(3, Col)
     rs("ContBrkPntOff") = .TextMatrix(4, Col)
     rs("OpeForceOff") = .TextMatrix(5, Col)
    'rs("ContResOff1") = .TextMatrix(6, Col)
     End With
    ElseIf Cavity = 2 Then
     With VSCavity2
          Col = 2
     rs("TestVolOn") = .TextMatrix(1, Col)
     rs("TestCurOn") = .TextMatrix(2, Col)
     rs("MVDOn") = .TextMatrix(3, Col)
     rs("ContBrkPntOn") = .TextMatrix(4, Col)
     rs("OpeForceOn") = .TextMatrix(5, Col)
     'rs("ContResOn1") = .TextMatrix(6, Col)
     Col = 1
     rs("TestVolOff") = .TextMatrix(1, Col)
     rs("TestCurOff") = .TextMatrix(2, Col)
     rs("MVDOff") = .TextMatrix(3, Col)
     rs("ContBrkPntOff") = .TextMatrix(4, Col)
     rs("OpeForceOff") = .TextMatrix(5, Col)
    'rs("ContResOff1") = .TextMatrix(6, Col)
     End With
    ElseIf Cavity = 3 Then
     With VSCavity3
          Col = 2
     rs("TestVolOn") = .TextMatrix(1, Col)
     rs("TestCurOn") = .TextMatrix(2, Col)
     rs("MVDOn") = .TextMatrix(3, Col)
     rs("ContBrkPntOn") = .TextMatrix(4, Col)
     rs("OpeForceOn") = .TextMatrix(5, Col)
     'rs("ContResOn1") = .TextMatrix(6, Col)
     Col = 1
     rs("TestVolOff") = .TextMatrix(1, Col)
     rs("TestCurOff") = .TextMatrix(2, Col)
     rs("MVDOff") = .TextMatrix(3, Col)
     rs("ContBrkPntOff") = .TextMatrix(4, Col)
     rs("OpeForceOff") = .TextMatrix(5, Col)
    'rs("ContResOff1") = .TextMatrix(6, Col)
     End With
    ElseIf Cavity = 4 Then
     With VSCavity4
          Col = 2
     rs("TestVolOn") = .TextMatrix(1, Col)
     rs("TestCurOn") = .TextMatrix(2, Col)
     rs("MVDOn") = .TextMatrix(3, Col)
     rs("ContBrkPntOn") = .TextMatrix(4, Col)
     rs("OpeForceOn") = .TextMatrix(5, Col)
     'rs("ContResOn1") = .TextMatrix(6, Col)
     Col = 1
     rs("TestVolOff") = .TextMatrix(1, Col)
     rs("TestCurOff") = .TextMatrix(2, Col)
     rs("MVDOff") = .TextMatrix(3, Col)
     rs("ContBrkPntOff") = .TextMatrix(4, Col)
     rs("OpeForceOff") = .TextMatrix(5, Col)
    'rs("ContResOff1") = .TextMatrix(6, Col)
     End With
    End If
'    With VSCavity2
'    Col = 1
'    rs("TestVolOn2") = .TextMatrix(1, Col)
'    rs("TestCurOn2") = .TextMatrix(2, Col)
'    rs("MVDOn2") = .TextMatrix(3, Col)
'    rs("ContBrkPntOn2") = .TextMatrix(4, Col)
'    rs("OpeForceOn2") = .TextMatrix(5, Col)
'    'rs("ContResOn2") = .TextMatrix(6, Col)
'    Col = 2
'    rs("TestVolOff2") = .TextMatrix(1, Col)
'    rs("TestCurOff2") = .TextMatrix(2, Col)
'    rs("MVDOff2") = .TextMatrix(3, Col)
'    rs("ContBrkPntOff2") = .TextMatrix(4, Col)
'    rs("OpeForceOff2") = .TextMatrix(5, Col)
'    'rs("ContResOff2") = .TextMatrix(6, Col)
'    End With
'
'    With VSCavity3
'    Col = 1
'    rs("TestVolOn3") = .TextMatrix(1, Col)
'    rs("TestCurOn3") = .TextMatrix(2, Col)
'    rs("MVDOn3") = .TextMatrix(3, Col)
'    rs("ContBrkPntOn3") = .TextMatrix(4, Col)
'    rs("OpeForceOn3") = .TextMatrix(5, Col)
'    'rs("ContResOn3") = .TextMatrix(6, Col)
'    Col = 2
'    rs("TestVolOff3") = .TextMatrix(1, Col)
'    rs("TestCurOff3") = .TextMatrix(2, Col)
'    rs("MVDOff3") = .TextMatrix(3, Col)
'    rs("ContBrkPntOff3") = .TextMatrix(4, Col)
'    rs("OpeForceOff3") = .TextMatrix(5, Col)
'    'rs("ContResOff3") = .TextMatrix(6, Col)
'    End With
'
'    With VSCavity4
'    Col = 1
'    rs("TestVolOn4") = .TextMatrix(1, Col)
'    rs("TestCurOn4") = .TextMatrix(2, Col)
'    rs("MVDOn4") = .TextMatrix(3, Col)
'    rs("ContBrkPntOn4") = .TextMatrix(4, Col)
'    rs("OpeForceOn4") = .TextMatrix(5, Col)
'    'rs("ContResOn4") = .TextMatrix(6, Col)
'    Col = 2
'    rs("TestVolOff4") = .TextMatrix(1, Col)
'    rs("TestCurOff4") = .TextMatrix(2, Col)
'    rs("MVDOff4") = .TextMatrix(3, Col)
'    rs("ContBrkPntOff4") = .TextMatrix(4, Col)
'    rs("OpeForceOff4") = .TextMatrix(5, Col)
'    'rs("ContResOff3") = .TextMatrix(6, Col)
'    End With
    
   rs.Update
End Sub
Private Sub SaveCounter()
Dim Sql As String
Dim rs As ADODB.Recordset
    A = Format(runningreportdate, "MM-dd-yyyy")
    Sql = "Select * from Model_Report_Counter  where Datetime = #" & Format(runningreportdate, "MM-dd-yyyy") & "# and shifttime = '" & Val(runningreportshift) & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    If rs.EOF = True Then
      rs.AddNew
      rs("ModelName") = ModelName
      rs("DateTime") = runningreportdate
      rs("ShiftTime") = runningreportshift
      rs("Mailsent") = 0
      rs("ModelNo") = ModelNo
    End If
      rs("ProductionCounter") = Val(txtproductioncounter.Text)
      'rs("OKCounter") = Val(txtOKCounter.Text)
      'rs("NGCounter") = Val(txtNGCounter.Text)
      'rs("CouplerCounter") = Val(txtCouplerCounter(0).Text)
      rs("BatchCounter") = Val(txtBatchCounter.Text)
      If Val(txtTargetProduction.Text) > 0 Then
        rs("TargetProduction") = Val(txtTargetProduction.Text)
      End If
      rs.Update
End Sub
Private Sub SaveBreakDown(breakdownType As Integer, breakdownstatus As Integer)
Dim Sql As String
Dim rs As ADODB.Recordset
   Sql = "Select Top 1 * from Model_Report_Breakdown "
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If breakdownstatus = 1 Then
      rs.AddNew
      rs("StartTime") = Format(Now(), "mm/dd/yyyy hh:mm:ss")
      rs("BreakdownType") = breakdownType
   Else
      rs("Remarks") = txtbreakdownsummary.Text
      rs("EndTime") = Format(Now(), "mm/dd/yyyy hh:mm:ss")
   End If
   rs.Update
   Exit Sub
Error:
   ErrorLog Err.Number, Err.Description, Erl, Me.Name, "SaveReport"
   Resume Next
End Sub

Private Sub SaveCounterValue()
 Dim ProdDay As String
 SaveSetting App.Title, ModelName, "OkCounter1", Val(txtOKCounter(0).Text)
 SaveSetting App.Title, ModelName, "NGCounter1", Val(txtNGCounter(0).Text)
 SaveSetting App.Title, ModelName, "OkCounter2", Val(txtOKCounter(1).Text)
 SaveSetting App.Title, ModelName, "NGCounter2", Val(txtNGCounter(1).Text)
 SaveSetting App.Title, ModelName, "OkCounter3", Val(txtOKCounter(2).Text)
 SaveSetting App.Title, ModelName, "NGCounter3", Val(txtNGCounter(2).Text)
 SaveSetting App.Title, ModelName, "OkCounter4", Val(txtOKCounter(3).Text)
 SaveSetting App.Title, ModelName, "NGCounter4", Val(txtNGCounter(3).Text)
 SaveSetting App.Title, ModelName, "CouplerCounter1", Val(txtCouplerCounter(0).Text)
 SaveSetting App.Title, ModelName, "CouplerCounter2", Val(txtCouplerCounter(1).Text)
 SaveSetting App.Title, ModelName, "CouplerCounter3", Val(txtCouplerCounter(2).Text)
 SaveSetting App.Title, ModelName, "CouplerCounter4", Val(txtCouplerCounter(3).Text)
 SaveSetting App.Title, ModelName, "BatchCounter", Val(txtBatchCounter.Text)
SaveSetting App.Title, ModelName, "TargetProduction", txtTargetProduction.Text
       
 'ProdDay = Format(Date, "ddmmyy")
 'SaveSetting App.Title, ModelName, "", Val(ProdDay)
 'SaveSetting App.Title, ModelName, "PrintCounter", txtprintcounter.Text
End Sub
Private Sub SaveProductioncounter()
Dim rs As ADODB.Recordset
Dim Sql As String
    Sql = "Select * from Model_Set where ModelName ='" & ModelName & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    rs("productioncounter") = Val(txtproductioncounter.Text)
    rs.Update
    'txtSaveCoupler.Text = Rs("CouplerCounter")
End Sub
Private Sub GetCounterValue()
On Error Resume Next
Dim ProdDay As String
Dim Today As String
   txtOKCounter(0).Text = Val(GetSetting(App.Title, ModelName, "OkCounter1", 0))
   txtNGCounter(0).Text = Val(GetSetting(App.Title, ModelName, "NgCounter1", 0))
   txtOKCounter(1).Text = Val(GetSetting(App.Title, ModelName, "OkCounter2", 0))
   txtNGCounter(1).Text = Val(GetSetting(App.Title, ModelName, "NgCounter2", 0))
   txtOKCounter(2).Text = Val(GetSetting(App.Title, ModelName, "OkCounter3", 0))
   txtNGCounter(2).Text = Val(GetSetting(App.Title, ModelName, "NgCounter3", 0))
   txtOKCounter(3).Text = Val(GetSetting(App.Title, ModelName, "OkCounter4", 0))
   txtNGCounter(3).Text = Val(GetSetting(App.Title, ModelName, "NgCounter4", 0))
   txtCouplerCounter(0).Text = Val(GetSetting(App.Title, ModelName, "CouplerCounter1", 0))
   txtCouplerCounter(1).Text = Val(GetSetting(App.Title, ModelName, "CouplerCounter2", 0))
   txtCouplerCounter(2).Text = Val(GetSetting(App.Title, ModelName, "CouplerCounter3", 0))
   txtCouplerCounter(3).Text = Val(GetSetting(App.Title, ModelName, "CouplerCounter4", 0))
   txtBatchCounter.Text = Val(GetSetting(App.Title, ModelName, "BatchCounter", 0))
   txtTargetProduction.Text = GetSetting(App.Title, ModelName, "TargetProduction", 0)
         
   tempshift = GetSetting(App.Title, ModelName, "saveshift", 0)
   TempDate = GetSetting(App.Title, ModelName, "savedate", 0)
   If TempDate <> runningreportdate Then
      txtOKCounter(0).Text = 0
      txtNGCounter(0).Text = 0
      txtOKCounter(1).Text = 0
      txtNGCounter(1).Text = 0
      txtOKCounter(2).Text = 0
      txtNGCounter(2).Text = 0
      txtOKCounter(3).Text = 0
      txtNGCounter(3).Text = 0
      SaveSetting App.Title, ModelName, "saveshift", runningreportshift
      SaveSetting App.Title, ModelName, "savedate", runningreportdate
      'txtprintcounter.Text = 0
   End If
   SaveCounterValue
End Sub

Private Function cmdCon()
   WinSock1.Close
   WinSock1.RemoteHost = IP_HOST
   WinSock1.RemotePort = IP_PORT
   WinSock1.Connect
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
   If (WinSock1.State = 7) And (CommandOn = False) Then
      Timer1.Enabled = False
      Select Case CommandType
         Case 1
            Call GetReadArray(StdReadStartAddress, StdReadCount, ReadArray)
            WinSock1.SendData ReadArray
            CVRead = CVRead + 1
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case 2
            Call GetWriteArray(StdWriteStartAddress, StdWriteCount, WriteArray)
            WinSock1.SendData WriteArray
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case 3
            Call GetReadArray((ExtendedReadStartAddress + (ExtendedReadCount * CVExtPktNo)), ExtendedReadCount, ReadArray)
            WinSock1.SendData ReadArray
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

   If (WinSock1.State <> 7) Then 'And (WinSock1.State <> 6) Then
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

Private Sub Timer8_Timer()
 If shapeInternet.BackColor = vbGreen Then
  sendEmail
 End If
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtBarcode.Locked = True
   If txtBarcode.Text = barcode Then
     txtBarcode.BackColor = vbGreen
     PLcdata(350) = 1
   Else
     txtBarcode.BackColor = vbRed
     PLcdata(350) = 2
     'SaveReport "NG"
   End If
End If
End Sub

Private Function ValidateBarcode(barcode As String) As Boolean
Dim rs As ADODB.Recordset
Dim Sql As String
   Sql = "Select * from Model_Report where barcode='" & barcode & "'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      checkBarcoderepeat = True
   Else
      checkBarcoderepeat = False
   End If
End Function



Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim SocketData() As Byte
Dim RegData, A, B, c As String
Dim i, j, K, l, M, n, ExpectedArraySize, ExtndedReadFrom, ExpectedLength As Integer
Dim Idata As Long
Dim Idata1 As Long

   Timer5.Enabled = False
   PLC_Communication_Error = False
   WinSock1.GetData SocketData
   CommandOn = False
   PlcCommCheck = False
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
