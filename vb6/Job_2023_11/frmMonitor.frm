VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMonitor 
   Caption         =   "MI_7646_USB_Charger"
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
      Height          =   9015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   18015
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   11
         Left            =   12720
         TabIndex        =   76
         Text            =   "0.000"
         Top             =   6300
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   10
         Left            =   12720
         TabIndex        =   73
         Text            =   "0.000"
         Top             =   5220
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   9
         Left            =   12720
         TabIndex        =   70
         Text            =   "0.000"
         Top             =   2820
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   8
         Left            =   12720
         TabIndex        =   67
         Text            =   "0.000"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   7
         Left            =   7560
         TabIndex        =   64
         Text            =   "0.000"
         Top             =   6300
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   6
         Left            =   7560
         TabIndex        =   61
         Text            =   "0.000"
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   5
         Left            =   7560
         TabIndex        =   58
         Text            =   "0.000"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   4
         Left            =   7560
         TabIndex        =   55
         Text            =   "0.000"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   3
         Left            =   2400
         TabIndex        =   52
         Text            =   "0.000"
         Top             =   6360
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   2400
         TabIndex        =   49
         Text            =   "0.000"
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   2400
         TabIndex        =   46
         Text            =   "0.000"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Online Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   15000
         TabIndex        =   40
         Top             =   1680
         Width           =   2895
         Begin VB.TextBox txtOd2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   360
            Width           =   990
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Width           =   735
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
         Height          =   390
         Left            =   5520
         TabIndex        =   34
         Top             =   720
         Width           =   7215
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
         Left            =   16440
         TabIndex        =   26
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
            Picture         =   "frmMonitor.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   27
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
         Height          =   2055
         Left            =   15000
         TabIndex        =   23
         Top             =   5880
         Width           =   2895
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
            Left            =   240
            TabIndex        =   25
            Top             =   240
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
            Left            =   300
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   15000
         TabIndex        =   18
         Top             =   2880
         Width           =   2895
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   2280
            Width           =   1290
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
            Height          =   390
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1320
            Width           =   1350
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   840
            Width           =   1350
         End
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
            Height          =   390
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   240
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox txtCouplerCounter 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   720
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Production Count"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   43
            Left            =   120
            TabIndex        =   80
            Top             =   2280
            Width           =   1020
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label8 
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
            TabIndex        =   33
            Top             =   1440
            Width           =   1335
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
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1335
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
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
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
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
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
         Left            =   15480
         TabIndex        =   15
         Top             =   240
         Width           =   2415
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
            TabIndex        =   29
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
            TabIndex        =   35
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
            TabIndex        =   28
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
            TabIndex        =   17
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
            TabIndex        =   16
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
         Height          =   630
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmMonitor.frx":0C42
         Top             =   8280
         Width           =   16215
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
         Left            =   10440
         TabIndex        =   11
         Top             =   -360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.Timer Timer8 
            Interval        =   60000
            Left            =   1440
            Top             =   1440
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
            TabIndex        =   12
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
            TabIndex        =   13
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
         Height          =   1140
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "MODEL DESC"
         Top             =   3600
         Width           =   14535
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   10440
         TabIndex        =   1
         Top             =   7320
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   5415
            Begin VB.TextBox txtIP 
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
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   6
               Text            =   "127.0.0.1"
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox txtPort 
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
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   5
               Text            =   "1232"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtIP_Host 
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
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   4
               Text            =   "127.0.0.1"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text1 
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
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   3
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label1 
               Caption         =   "IP M/C"
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
               Index           =   4
               Left            =   240
               TabIndex        =   9
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label16 
               Caption         =   "PORT:"
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
               Left            =   2520
               TabIndex        =   8
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label15 
               Caption         =   "IP Host"
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
               Left            =   1440
               TabIndex        =   7
               Top             =   240
               Width           =   615
            End
         End
      End
      Begin VB.TextBox txtLedCurrent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   2400
         TabIndex        =   44
         Text            =   "0.000"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   10920
         TabIndex        =   78
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   14520
         TabIndex        =   77
         Top             =   6360
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   10920
         TabIndex        =   75
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   14520
         TabIndex        =   74
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   10920
         TabIndex        =   72
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   14520
         TabIndex        =   71
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   10920
         TabIndex        =   69
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   14520
         TabIndex        =   68
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   5760
         TabIndex        =   66
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   9360
         TabIndex        =   65
         Top             =   6360
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   5760
         TabIndex        =   63
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   9360
         TabIndex        =   62
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   5760
         TabIndex        =   60
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9360
         TabIndex        =   59
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   57
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   9360
         TabIndex        =   56
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   54
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   53
         Top             =   6480
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   51
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   50
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   48
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   47
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblAmp 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   45
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblLed 
         Caption         =   "LED - 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   43
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   2640
         TabIndex        =   36
         Top             =   840
         Width           =   1275
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
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   9480
      TabIndex        =   39
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
      TabIndex        =   38
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
      TabIndex        =   37
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
    txtNGCounter.Text = 0
    SaveCounterValue
  End If
End Sub

Private Sub CmdOKCounter_Click()
If MsgBox("Are you Sure You Want To Reset OK Counter", vbInformation + vbYesNo) = vbYes Then
    txtOKCounter.Text = 0
    SaveCounterValue
  End If
End Sub

Private Sub cmdclosebreakdownscreen_Click()
    PictureBreakdown.Visible = False
    Command2.Enabled = True
End Sub

Private Sub cmdfullbreakdown_Click()
    cmdrunningbreakdown.Enabled = False
    cmdfullbreakdown.Enabled = False
    cmdgolive.Enabled = True
    cmdclosebreakdownscreen.Enabled = False
    SaveBreakDown 3, 1
    PLcdata(248) = 3
End Sub

Private Sub cmdgolive_Click()
    cmdrunningbreakdown.Enabled = True
    cmdfullbreakdown.Enabled = True
    cmdgolive.Enabled = False
    cmdclosebreakdownscreen.Enabled = True
    SaveBreakDown 1, 0
    PLcdata(248) = 1
End Sub

Private Sub cmdrunningbreakdown_Click()
    cmdrunningbreakdown.Enabled = False
    cmdfullbreakdown.Enabled = False
    cmdgolive.Enabled = True
    cmdclosebreakdownscreen.Enabled = False
    SaveBreakDown 2, 1
    PLcdata(248) = 2
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
      PLcdata(249) = 0
  Else
    txtTargetProduction.BackColor = vbRed
  End If
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    PictureBreakdown.Visible = True
End Sub


Private Sub Command3_Click()
'PLcdata(109) = 1
'AssignPLCdata
'sendEmail
SaveReport (1)
'PrintLabel JustPrinter1
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
   Winsock1.Protocol = sckTCPProtocol
   txtIP.Text = Winsock1.LocalIP
   txtIP_Host = rs("PLC_IP") '"192.168.1.30"
   txtPort = rs("PLC_Port")
   rs.Close
Exit Sub
Error:
rs.Close
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
Frame1.Top = ((Screen.Height - Frame1.Height) / 2) - 100
Frame1.Left = ((Screen.Width - Frame1.Width) / 2)
LoadSettingsData
Call Load_Message_File
runningreportshift = GetSetting(App.Title, ModelName, "saveshift", 0)
runningreportdate = GetSetting(App.Title, ModelName, "savedate", 0)
PLcdata(240) = 1
GetCounterValue
ConnectToPLC
LoadGrid
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
      'CmdOKCounter.Visible = False
      'CmdNgCounter.Visible = False
      'Command1.Visible = False
   ElseIf AccessType = "1" Then 'Disable or Hide for AccessType 1
      'CmdOKCounter.Visible = False
      'CmdNgCounter.Visible = False
      'Command1.Visible = False
   ElseIf AccessType = "2" Then 'Show All Which Will Disable or Hide For One
      'CmdOKCounter.Visible = True
      'CmdNgCounter.Visible = True
   End If
End Sub
Private Sub LoadGrid()
Dim rs As ADODB.Recordset
Dim Sql As String

   Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
With Grid1
    '.CellAlignment = flexAlignCenterCenter
    .RowHeight(0) = 1000
    .ColWidthMin = 1100
    .ColWidthMax = 1200
    .Cols = 6
    .Rows = 2
    .TextMatrix(0, 0) = "Reverse" & vbNewLine & "Polarity"
    .TextMatrix(0, 1) = "Cut-Off" & vbNewLine & "Voltage" & vbNewLine & "(AT<" & rs("CutoffVolt") & "V)"
    .TextMatrix(0, 2) = "O/P" & vbNewLine & "Voltage" & vbNewLine & "(AT " & rs("OutputVolt1") & "V)"
    .TextMatrix(0, 3) = "O/P" & vbNewLine & "Voltage" & vbNewLine & "(AT " & rs("OutputVolt2") & "V)"
    .TextMatrix(0, 4) = "O/P" & vbNewLine & "Voltage" & vbNewLine & "(AT " & rs("OutputVolt3") & "V)"
    .TextMatrix(0, 5) = "O/P" & vbNewLine & "Voltage" & vbNewLine & "Short Test"
    If rs("bypass9").Value = 1 Then
     .ColHidden(0) = True
    End If
    If rs("bypass10").Value = 1 Then
     .ColHidden(1) = True
    End If
    If rs("bypass11").Value = 1 Then
     .ColHidden(2) = True
    End If
    If rs("bypass12").Value = 1 Then
     .ColHidden(3) = True
    End If
    If rs("bypass13").Value = 1 Then
     .ColHidden(4) = True
    End If
    If rs("bypass14").Value = 1 Then
     .ColHidden(5) = True
    End If
End With
With Grid2
    '.CellAlignment = flexAlignLeftCenter
    .RowHeight(0) = 1000
    .ColWidthMin = 1100
    .ColWidthMax = 1200
    .Cols = 6
    .Rows = 2
    .TextMatrix(0, 0) = "Input" & vbNewLine & "Voltage"
    .TextMatrix(0, 1) = "Input" & vbNewLine & "Current"
    .TextMatrix(0, 2) = "O/P" & vbNewLine & "Voltage"
    .TextMatrix(0, 3) = "O/P" & vbNewLine & "Current"
    .TextMatrix(0, 4) = "Efficiency" & vbNewLine & "(>75%)"
    .TextMatrix(0, 5) = "Result"
End With
    rs.Close
End Sub
'Private Sub Chartload()
'A = MSChart1.Plot.SeriesCollection.Count
'
'MSChart1.Plot.SeriesCollection(1).LegendText = "Input"
'MSChart1.Plot.SeriesCollection(2).LegendText = "Output"
''mschart1.Plot.
'End Sub
Private Function AssignPLCdata()
 On Error GoTo Error
   MsgCode = PLcdata(108)
   GridColorfunction1 Grid1, 1, 0, PLcdata(100), &H1, &H2
   GridColorfunction Grid1, 1, 1, PLcdata(100), &H4, &H8
   GridColorfunction Grid1, 1, 2, PLcdata(100), &H10, &H20
   GridColorfunction Grid1, 1, 3, PLcdata(100), &H40, &H80
   GridColorfunction Grid1, 1, 4, PLcdata(100), &H100, &H200
   GridColorfunction1 Grid1, 1, 5, PLcdata(100), &H400, &H800
   
   GridColorfunction Grid2, 1, 0, PLcdata(101), &H1, &H2
   GridColorfunction Grid2, 1, 1, PLcdata(101), &H4, &H8
   GridColorfunction Grid2, 1, 2, PLcdata(101), &H10, &H20
   GridColorfunction Grid2, 1, 3, PLcdata(101), &H40, &H80
   GridColorfunction Grid2, 1, 4, PLcdata(101), &H100, &H200
   GridColorfunction1 Grid2, 1, 5, PLcdata(101), &H400, &H800

   GridTextFunction Grid1, 1, 1, PLcdata(110), 100, "0.00"
   GridTextFunction Grid1, 1, 2, PLcdata(111), 100, "0.00"
   GridTextFunction Grid1, 1, 3, PLcdata(112), 100, "0.00"
   GridTextFunction Grid1, 1, 4, PLcdata(113), 100, "0.00"

   GridTextFunction Grid2, 1, 0, PLcdata(120), 100, "0.00"
   GridTextFunction Grid2, 1, 1, PLcdata(121), 1000, "0.000"
   GridTextFunction Grid2, 1, 2, PLcdata(122), 100, "0.00"
   GridTextFunction Grid2, 1, 3, PLcdata(123), 1000, "0.000"
   GridTextFunction Grid2, 1, 4, PLcdata(124), 1, "00"

   TextColorfunction txtResistance1(0), PLcdata(106), &H1, &H2
   TextColorfunction txtResistance2(0), PLcdata(106), &H4, &H8
   TextColorfunction txtResistance1(1), PLcdata(106), &H10, &H20
   TextColorfunction txtResistance2(1), PLcdata(106), &H40, &H80
   
   txtOD1.Text = Format(PLcdata(102) / 100, "0.00")
   txtOd2.Text = Format(PLcdata(103) / 1000, "0.000")
   txtOD3.Text = Format(PLcdata(104) / 100, "0.00")
   txtOD4.Text = Format(PLcdata(105) / 1000, "0.000")
   
   txtCycleTime.Text = Format(PLcdata(107) / 10, "0.0")

   txtResistance1(0).Text = Format(PLcdata(130) / 1000, "0.000")
   txtResistance2(0).Text = Format(PLcdata(131), "0")
   txtResistance1(1).Text = Format(PLcdata(132) / 1000, "0.000")
   txtResistance2(1).Text = Format(PLcdata(133), "0")
   
   If PLcdata(165) = 0 And pulseBreakdown = True Then
      pulseBreakdown = False
      'PictureBreakdown.Visible = False
   ElseIf PLcdata(165) = 1 And pulseBreakdown = False Then
      pulseBreakdown = True
      PictureBreakdown.Visible = True
      cmdrunningbreakdown.Enabled = False
      cmdfullbreakdown.Enabled = False
      cmdgolive.Enabled = True
      cmdclosebreakdownscreen.Enabled = False
      
   ElseIf PLcdata(165) = 2 And pulseBreakdown = False Then
      pulseBreakdown = True
      PictureBreakdown.Visible = True
       cmdrunningbreakdown.Enabled = False
      cmdfullbreakdown.Enabled = False
      cmdgolive.Enabled = True
      cmdclosebreakdownscreen.Enabled = False
   End If
   
   If PLcdata(170) = 0 And PulseScan = False Then
      PulseScan = True
      txtBarcode.Locked = False
      txtBarcode.BackColor = vbWhite
      txtBarcode.Text = ""
      txtBarcode.Locked = True
      PLcdata(250) = 0
   ElseIf PLcdata(170) = 1 And PulseScan = True Then
      PulseScan = False
      txtBarcode.Locked = False
      txtBarcode.BackColor = vbYellow
      txtBarcode.Text = ""
      txtBarcode.SetFocus
      
   End If
   If PLcdata(109) = 0 And pulse1 = False Then
      pulse1 = True
      lblGo.Visible = False
      lblNg.Visible = False
   ElseIf PLcdata(109) = 1 And pulse1 = True Then
      pulse1 = False
      lblGo.Visible = True
      GetCounterValue
      txtProductionCounter.Text = Val(txtProductionCounter.Text) + 1
      txtOKCounter.Text = Val(txtOKCounter.Text) + 1
      txtBatchCounter.Text = Val(txtBatchCounter.Text) + 1
      txtTargetProduction.Text = Val(txtTargetProduction.Text) - 1
      txtCouplerCounter.Text = Val(txtCouplerCounter.Text) + 1
      If pulsePrinterBypass = False Then
        PrintLabel JustPrinter1
      End If
      SaveProductioncounter
      SaveReport 1
      SaveCounter
      SaveCounterValue
   ElseIf PLcdata(109) = 2 And pulse1 = True Then
      pulse1 = False
      GetCounterValue
      lblNg.Visible = True
      txtNGCounter.Text = Val(txtNGCounter.Text) + 1
      SaveReport 2
      SaveCounter
      SaveCounterValue
   End If
      
Exit Function
Error:
   ErrorLog Err.Number, Err.Description & "---", Erl, Me.Name, "Assign PLC Data"
   Resume Next
End Function

Private Sub GridTextFunction(Grid As VSFlexGrid, Row As Integer, Col As Integer, data As Integer, Devision As Integer, formatstring As String)
Grid.TextMatrix(Row, Col) = Format(data / Devision, formatstring)
End Sub
Private Sub GridColorfunction(Grid As VSFlexGrid, Row As Integer, Col As Integer, data As Integer, reg1 As Integer, reg2 As Integer)
    If (data And reg1) Then
        If (data And reg2) Then
           Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbYellow
           If Grid.TextMatrix(Row, Col) = "" Then
            Grid.TextMatrix(Row, Col) = "Testing"
           End If
        Else
           Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbGreen
           If Grid.TextMatrix(Row, Col) = "" Then
            Grid.TextMatrix(Row, Col) = "OK"
           End If
        End If
    ElseIf (data And reg2) Then
          Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
          If Grid.TextMatrix(Row, Col) = "" Then
            Grid.TextMatrix(Row, Col) = "NG"
          End If
    Else
          Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbWhite
          
    End If
End Sub

Private Sub TextColorfunction(ctrl As TextBox, data As Integer, reg1 As Integer, reg2 As Integer)
    If (data And reg1) Then
        If (data And reg2) Then
           ctrl.BackColor = vbYellow
        Else
           ctrl.BackColor = vbGreen
        End If
    ElseIf (data And reg2) Then
          ctrl.BackColor = vbRed
    Else
          ctrl.BackColor = vbWhite
    End If
    
End Sub

Private Sub GridColorfunction1(Grid As VSFlexGrid, Row As Integer, Col As Integer, data As Integer, reg1 As Integer, reg2 As Integer)
    If (data And reg1) Then
        If (data And reg2) Then
           Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbYellow
            Grid.TextMatrix(Row, Col) = "Testing"
        Else
           Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbGreen
           Grid.TextMatrix(Row, Col) = "OK"
        End If
    ElseIf (data And reg2) Then
          Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
            Grid.TextMatrix(Row, Col) = "NG"
    Else
          Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbWhite
          Grid.TextMatrix(Row, Col) = ""
          
    End If
End Sub
Private Sub ShapeColorsinglefunction(data As Integer, reg1 As Integer, ctrl As Object)
    If (data And reg1) <> 0 Then
          ctrl.BackColor = vbYellow
    Else
          ctrl.BackColor = vbWhite
    End If
End Sub
Private Sub ShapeColorsingleifunction(data As Integer, reg1 As Integer, ctrl As Object)
    If (data And reg1) <> 0 Then
          ctrl.BackColor = vbGreen
    Else
          ctrl.BackColor = vbWhite
    End If
End Sub

Private Sub Timer2_Timer()
'On Error GoTo Error

'    txttime = Format(Time(), "Hh:Mm:Ss")
'Chartload
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
    
    Text1.Text = WinsockStstus(Winsock1.State)


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
'End Sub
'Error:
'ErrorLog Err.Number, Err.Description, Erl, Me.Name, "Timer2"
End Sub
Public Function sendEmail()
On Error GoTo Error
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
Exit Function
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "SendEmail"
End Function
Private Function callSendEmailApi(rsGeneralset As ADODB.Recordset, Subject As String, Body As String) As Boolean
On Error GoTo Error
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
     If response = "SENT" Then
     callSendEmailApi = True
     Else
     callSendEmailApi = False
     
     End If
     
    
    End With
Exit Function
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "CallsendEmailApi"
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

Private Sub LoadData()

On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
Dim strByPass(14) As Integer
Dim j As Integer

    Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    PLcdata(240) = 1
    PLcdata(203) = Val(rs("CutoffVolt")) * 100
    PLcdata(204) = Val(rs("OutputVolt1")) * 100
    PLcdata(205) = Val(rs("OutputVolt2")) * 100
    PLcdata(206) = Val(rs("OutputVolt3")) * 100
    PLcdata(207) = Val(rs("testVoltage")) * 100
    PLcdata(208) = Val(rs("testCurrent"))
    PLcdata(209) = Val(rs("EfficiencyMin"))
    PLcdata(210) = Val(rs("EfficiencyMax"))
    PLcdata(211) = Val(rs("InputCurrentMin")) * 1000
    PLcdata(212) = Val(rs("InputCurrentMax")) * 1000
    PLcdata(213) = Val(rs("OutputVoltMin")) * 100
    PLcdata(214) = Val(rs("OutputVoltMax")) * 100
    PLcdata(215) = Val(rs("OutputCurrentMin")) * 1000
    PLcdata(216) = Val(rs("OutputCurrentMax")) * 1000
    PLcdata(217) = Val(rs("VoltageOffset")) * 100
    PLcdata(218) = Val(rs("CurrentOffset")) * 1000
    PLcdata(219) = Val(rs("Efficiencyoffset")) * 100
    
        PLcdata(245) = 0

         PLcdata(220) = Val(rs("CutoffVoltMin")) * 100
         PLcdata(221) = Val(rs("CutoffVoltMax")) * 100
         PLcdata(222) = Val(rs("OutputVolt1Min")) * 100
         PLcdata(223) = Val(rs("OutputVolt1Max")) * 100
         PLcdata(224) = Val(rs("OutputVolt2Min")) * 100
         PLcdata(225) = Val(rs("OutputVolt2Max")) * 100
         PLcdata(226) = Val(rs("OutputVolt3Min")) * 100
         PLcdata(227) = Val(rs("OutputVolt3Max")) * 100
         
    PLcdata(228) = Val(rs("InputVoltageOffset")) * 100
    PLcdata(229) = Val(rs("InputCurrentOffset")) * 1000
    
    PLcdata(260) = Val(rs("Resistance1Offset")) * 1000
    PLcdata(261) = Val(rs("Resistance2Offset"))
   
    PLcdata(265) = Val(rs("Resistance1Min")) * 1000
    PLcdata(266) = Val(rs("Resistance1Max")) * 1000
    PLcdata(267) = Val(rs("Resistance2Min"))
    PLcdata(268) = Val(rs("Resistance2Max"))
    
    txtModelDesc.Text = Trim(rs("ModelDesc"))
    If Val(txtCouplerCounter.Text) >= setCouplerCounter Then
        PLcdata(235) = 1
    ElseIf Val(txtBatchCounter.Text) >= setBatchCounter Then
        PLcdata(235) = 2
    Else
        PLcdata(235) = 0
    End If
    
    PartNo = rs("PrintPartNo")
'    BarcodeLength = rs("BarcodeLength")
    HardwareNo = rs("HardwareNo")
'    SerialStartingtxt = rs("SerialStartingtxt")
    
    PLcdata(232) = Val(rs("DotMarkingTime")) * 10

    ModelNo = rs("ModelNo")
    PLcdata(233) = rs("ModelNo")
    'PLcdata(234) = Val(rs("ScanDelayTime")) * 10
    PLcdata(230) = 0
    PLcdata(230) = PLcdata(230) + &H1 * Val(rs("Bypass1"))
    PLcdata(230) = PLcdata(230) + &H2 * Val(rs("Bypass2"))
    PLcdata(230) = PLcdata(230) + &H4 * Val(rs("Bypass3"))
    PLcdata(230) = PLcdata(230) + &H8 * Val(rs("Bypass4"))
    PLcdata(230) = PLcdata(230) + &H10 * Val(rs("Bypass5"))
    PLcdata(230) = PLcdata(230) + &H20 * Val(rs("Bypass6"))
    PLcdata(230) = PLcdata(230) + &H40 * Val(rs("Bypass7"))
    PLcdata(230) = PLcdata(230) + &H80 * Val(rs("ByPass8"))
    PLcdata(230) = PLcdata(230) + &H100 * Val(rs("Bypass15"))
    PLcdata(230) = PLcdata(230) + &H200 * Val(rs("Bypass16"))
    PLcdata(230) = PLcdata(230) + &H400 * Val(rs("Bypass17"))
    PLcdata(230) = PLcdata(230) + &H800 * Val(rs("Bypass18"))
    PLcdata(230) = PLcdata(230) + &H1000 * Val(rs("Bypass19"))
    
    PLcdata(231) = 0
    PLcdata(231) = PLcdata(231) + &H1 * Val(rs("Bypass9"))
    PLcdata(231) = PLcdata(231) + &H2 * Val(rs("Bypass10"))
    PLcdata(231) = PLcdata(231) + &H4 * Val(rs("Bypass11"))
    PLcdata(231) = PLcdata(231) + &H8 * Val(rs("Bypass12"))
    PLcdata(231) = PLcdata(231) + &H10 * Val(rs("Bypass13"))
    PLcdata(231) = PLcdata(231) + &H20 * Val(rs("Bypass14"))

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
            PLcdata(249) = 1
            Exit Sub
        Else
            If tempgetshift <> tempshift Then
                txtTargetProduction.Locked = False
                txtTargetProduction.Text = ""
                txtTargetProduction.SetFocus
                txtTargetProduction.BackColor = vbYellow
                Command1.Visible = True
                PLcdata(249) = 1
                Exit Sub
            End If
        End If
    ElseIf txtTargetProduction.BackColor <> vbYellow Then
        txtTargetProduction.Locked = False
        txtTargetProduction.Text = ""
        txtTargetProduction.SetFocus
        txtTargetProduction.BackColor = vbYellow
        Command1.Visible = True
        PLcdata(249) = 1
        
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
    If Val(rs("Bypass1")) = 1 Then
        Grid2.Visible = False
        Label10(1).Visible = False
    End If
    
    If Val(rs("Bypass18")) = 1 Then
        txtResistance1(0).Visible = False
        txtResistance1(1).Visible = False
        Label19.Visible = False
        Label23.Visible = False
        Label25.Visible = False
        Label26.Visible = False
    End If
    
    If Val(rs("Bypass19")) = 1 Then
        txtResistance2(0).Visible = False
        txtResistance2(1).Visible = False
        Label20.Visible = False
        Label22.Visible = False
        
        Label27.Visible = False
        Label28.Visible = False
    End If
    If Val(rs("Bypass18")) = 1 And Val(rs("Bypass19")) = 1 Then
        Label21.Visible = False
        Label24.Visible = False
        Label25.Visible = False
        Label26.Visible = False
        Label27.Visible = False
        Label28.Visible = False
    End If
    PartNo = rs("PrintPartNo")
    'BarcodeLength = rs("BarcodeLength")
    HardwareNo = rs("HardwareNo")
    'SerialStartingtxt = rs("SerialStartingtxt")
    setBatchCounter = rs("BatchCounter")
    setCouplerCounter = rs("CouplerCounter")
    VendorId = rs("VandorId")
    ImgPart.Picture = LoadPicture(rs("PartImage"))
    txtProductionCounter.Text = rs("productioncounter")
    If Val(rs("PrinterBypass")) = 1 Then
        pulsePrinterBypass = True
    Else
        pulsePrinterBypass = False
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

Private Sub SaveReport(result As String)
On Error GoTo Error
Dim Sql As String
Dim rs As ADODB.Recordset
   Sql = "Select * from Model_Report"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   rs.AddNew
      rs("ModelName") = ModelName
      rs("OperatorName") = LoginUser
      rs("Date") = Format(printDateTime, "dd/mm/yyyy")
      rs("Time") = Format(printDateTime, "hh:mm:ss")
      rs("Barcode") = barcode
      rs("Result") = result
      rs("ClkResistance1Result") = getresultbycolorText(txtResistance1(0))
      rs("ClkResistance2Result") = getresultbycolorText(txtResistance2(0))
      rs("AClkResistance1Result") = getresultbycolorText(txtResistance1(1))
      rs("AClkResistance2Result") = getresultbycolorText(txtResistance2(1))
      
      rs("ClkResistance1") = txtResistance1(0).Text
      rs("ClkResistance2") = txtResistance2(0).Text
      rs("AClkResistance1") = txtResistance1(1).Text
      rs("AClkResistance2") = txtResistance2(1).Text
      
With Grid1
    rs("ReversePolarity") = .TextMatrix(1, 0)
    rs("CutOffVoltage") = .TextMatrix(1, 1)
    rs("Output1") = .TextMatrix(1, 2)
    rs("Output2") = .TextMatrix(1, 3)
    rs("Output3") = .TextMatrix(1, 4)
    rs("OutputShortTest") = .TextMatrix(1, 5)
    
    rs("CutOffVoltageStatus") = getresultbycolor(Grid1, 1, 1)
    rs("Output1Status") = getresultbycolor(Grid1, 1, 2)
    rs("Output2Status") = getresultbycolor(Grid1, 1, 3)
    rs("Output3Status") = getresultbycolor(Grid1, 1, 4)
End With
With Grid2
    rs("TestVoltage") = .TextMatrix(1, 0)
    rs("InputCurrent") = .TextMatrix(1, 1)
    rs("OPVoltage") = .TextMatrix(1, 2)
    rs("OPCurrent") = .TextMatrix(1, 3)
    rs("Efficiency") = .TextMatrix(1, 4)
    
    rs("TestVoltageStatus") = getresultbycolor(Grid2, 1, 0)
    rs("InputCurrentStatus") = getresultbycolor(Grid2, 1, 1)
    rs("OPVoltageStatus") = getresultbycolor(Grid2, 1, 2)
    rs("OPCurrentStatus") = getresultbycolor(Grid2, 1, 3)
    rs("EfficiencyStatus") = getresultbycolor(Grid2, 1, 4)
End With
    rs.Update
    Exit Sub
Error:
   ErrorLog Err.Number, Err.Description, Erl, Me.Name, "SaveReport"
   Resume Next
End Sub
Private Function getresultbycolor(Grid As VSFlexGrid, Row As Integer, Col As Integer) As String
 If Grid.Cell(flexcpBackColor) = vbRed Then
    getresultbycolor = "NG"
 ElseIf Grid.Cell(flexcpBackColor) = vbRed Then
    getresultbycolor = "OK"
 ElseIf Grid.Cell(flexcpBackColor) = vbRed Then
    getresultbycolor = "Testing"
 Else
    getresultbycolor = "-"
 End If
End Function
Private Function getresultbycolorText(ctrl As TextBox) As String
 If ctrl.BackColor = vbRed Then
    getresultbycolorText = "NG"
 ElseIf ctrl.BackColor = vbRed Then
    getresultbycolorText = "OK"
 ElseIf ctrl.BackColor = vbRed Then
    getresultbycolorText = "Testing"
 Else
    getresultbycolorText = "-"
 End If
End Function

Private Sub SaveCounter()
Dim Sql As String
Dim rs As ADODB.Recordset
    Sql = "Select * from Model_Report_Counter where datetime = #" & Format(runningreportdate, "MM-dd-yyyy") & "# and shifttime = '" & runningreportshift & "'"
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
      rs("ProductionCounter") = Val(txtProductionCounter.Text)
      rs("OKCounter") = Val(txtOKCounter.Text)
      rs("NGCounter") = Val(txtNGCounter.Text)
      rs("CouplerCounter") = Val(txtCouplerCounter.Text)
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
 SaveSetting App.Title, ModelName, "OkCounter", Val(txtOKCounter.Text)
 SaveSetting App.Title, ModelName, "NGCounter", Val(txtNGCounter.Text)
 SaveSetting App.Title, ModelName, "CouplerCounter", Val(txtCouplerCounter.Text)
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
    rs("productioncounter") = Val(txtProductionCounter.Text)
    rs.Update
    'txtSaveCoupler.Text = Rs("CouplerCounter")
End Sub
Private Sub GetCounterValue()
On Error Resume Next
Dim ProdDay As String
Dim Today As String
   txtOKCounter.Text = Val(GetSetting(App.Title, ModelName, "OkCounter", 0))
   txtNGCounter.Text = Val(GetSetting(App.Title, ModelName, "NgCounter", 0))
   txtCouplerCounter.Text = Val(GetSetting(App.Title, ModelName, "CouplerCounter", 0))
   txtBatchCounter.Text = Val(GetSetting(App.Title, ModelName, "BatchCounter", 0))
   txtTargetProduction.Text = GetSetting(App.Title, ModelName, "TargetProduction", 0)
         
   tempshift = GetSetting(App.Title, ModelName, "saveshift", 0)
   TempDate = GetSetting(App.Title, ModelName, "savedate", 0)
   If TempDate <> runningreportdate Or runningreportshift <> tempshift Then
      txtOKCounter.Text = 0
      txtNGCounter.Text = 0
      SaveSetting App.Title, ModelName, "saveshift", runningreportshift
      SaveSetting App.Title, ModelName, "savedate", runningreportdate
      'txtprintcounter.Text = 0
   End If
   If TempDate <> runningreportdate Then
        txtProductionCounter.Text = 0
        SaveCounter
   End If
   SaveCounterValue
End Sub

Private Function cmdCon()
   Winsock1.Close
   Winsock1.RemoteHost = txtIP_Host.Text
   Winsock1.RemotePort = txtPort.Text
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
     PLcdata(250) = 1
   Else
     txtBarcode.BackColor = vbRed
     PLcdata(250) = 2
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
   Winsock1.GetData SocketData
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
   text2 = CommandType & "+" & CVExtPktNo
   Timer1.Interval = 10
   Timer1.Enabled = True
End Sub
