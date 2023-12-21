VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMonitor 
   Caption         =   "Napino_BarcodeScan_2023_12"
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
      Height          =   9735
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   15855
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2640
         Width           =   4530
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2040
         Width           =   4530
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2760
         Width           =   4530
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2040
         Width           =   4530
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   4080
         Width           =   1890
      End
      Begin VB.Frame frmCoupler 
         BackColor       =   &H80000004&
         Caption         =   "Trey Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   0
         Left            =   12840
         TabIndex        =   18
         Top             =   3240
         Width           =   2535
         Begin VB.TextBox txtproductioncounter 
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
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2160
            Width           =   1830
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
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1200
            Width           =   615
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   720
            Width           =   990
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
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   480
            TabIndex        =   38
            Top             =   1800
            Width           =   1575
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
            TabIndex        =   21
            Top             =   360
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
            TabIndex        =   20
            Top             =   840
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
         TabIndex        =   17
         Text            =   "MODEL DESC"
         Top             =   240
         Width           =   12255
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
         TabIndex        =   13
         Text            =   "frmMonitor.frx":0000
         Top             =   8760
         Width           =   13695
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
         Left            =   13080
         TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   12
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
         Height          =   1815
         Left            =   12840
         TabIndex        =   5
         Top             =   6720
         Width           =   2535
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
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Visible         =   0   'False
            Width           =   2175
         End
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
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   2265
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
         Left            =   14040
         TabIndex        =   3
         Top             =   8760
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
            TabIndex        =   4
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
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtPartScanned 
         Alignment       =   2  'Center
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
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   6480
         Width           =   1890
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
         TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFModel 
         Height          =   5325
         Left            =   360
         TabIndex        =   26
         Top             =   3360
         Width           =   9675
         _cx             =   17066
         _cy             =   9393
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483638
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmMonitor.frx":0C54
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label lblBarcode 
         Caption         =   "Barcode - 4"
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
         Index           =   3
         Left            =   7440
         TabIndex        =   36
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label lblBarcode 
         Caption         =   "Barcode - 3"
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
         Index           =   2
         Left            =   7440
         TabIndex        =   34
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label lblBarcode 
         Caption         =   "Barcode - 2"
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
         Index           =   1
         Left            =   360
         TabIndex        =   32
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label lblBarcode 
         Caption         =   "Barcode - 1"
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
         Left            =   360
         TabIndex        =   30
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part in 1 Trey"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   10680
         TabIndex        =   28
         Top             =   3360
         Width           =   1680
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
         Index           =   7
         Left            =   6240
         TabIndex        =   24
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Scanned"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   43
         Left            =   10680
         TabIndex        =   19
         Top             =   5880
         Width           =   1710
      End
      Begin VB.Image ImgPart 
         Height          =   1815
         Left            =   17400
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   2055
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
Dim PartNumber As String
'----------------
Dim PLC_Communication_Error As Boolean
Dim MsgText() As String
Dim MsgColor() As Integer
Dim MsgCount As Integer
Dim CloseScreen As Boolean
Dim runningreportdate As Date
Dim runningreportshift As String
Dim ModelNo As Integer
Dim SqlBypass As Boolean
Dim BarcodeRepeatEnable As Boolean
Dim BarcodeValidation As Boolean
Dim PartNoValidation As Boolean
  
Private Sub cmdClose_Click()
CloseScreen = True
CloseMe
End Sub

Private Sub CloseMe()
If Not Con1 Is Nothing Then
If Con1.State = True Then Con1.Close
End If
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
CheckBarcodeValidation "121213442"
SaveReport 1
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
   ''ComPort(1) = rs("ComPort1")
''    ComPort(2) = Rs("ComPort2")
    ''ComPortBP(1) = rs("ComPortBP1")
''      ComPortBP(2) = Rs("ComPortBP2")
   ''PrinterName = rs("PrinterName1")
   SQLpath = rs("ConnString")
   Initialise
   Winsock1.Protocol = sckTCPProtocol
   'txtIP.Text = WinSock1.LocalIP
   IP_HOST = rs("PLC_IP") '"192.168.1.30"
   IP_PORT = rs("PLC_Port")
Exit Sub

Error:
 If Err.number = 8002 Then
    MsgBox "Com Port " & ComPort(Erl) & " Not Working", vbInformation
 ElseIf Err.number = 8005 Then
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
ConnectToPLC
LoadSettingsData
Call Load_Message_File
LoadGrid
PLcdata(340) = 1
GetCounterValue

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
On Error GoTo Error
  If AccessType = "0" Then 'Disable or Hide For Operator
   ElseIf AccessType = "1" Then 'Disable or Hide for AccessType 1
   ElseIf AccessType = "2" Then 'Show All Which Will Disable or Hide For One
   End If

Exit Sub
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "UserAccess"
   Resume Next
End Sub

Private Function AssignPLCdata()
On Error GoTo Error
   MsgCode = PLcdata(108)
   txtCycleTime.Text = Format(PLcdata(107) / 10, "0.0")
   If PLcdata(101) = 0 And pulse2 = False Then
        pulse2 = True
        PLcdata(201) = 0
        loadGridData
   ElseIf PLcdata(101) = 1 And pulse2 = True Then
        pulse2 = False
        If deletedummyData = 1 Then
            PLcdata(201) = 1
        End If
        loadGridData
   End If
   If PLcdata(103) = 0 And pulse3 = False Then
        pulse3 = True
        txtBarcode(0).Text = ""
        txtBarcode(1).Text = ""
        txtBarcode(2).Text = ""
        txtBarcode(3).Text = ""
        
        txtBarcode(0).BackColor = vbWhite
        txtBarcode(1).BackColor = vbWhite
        txtBarcode(2).BackColor = vbWhite
        txtBarcode(3).BackColor = vbWhite
        PLcdata(203) = 0
        loadGridData
   ElseIf PLcdata(103) = 1 And pulse3 = True Then
        pulse3 = False
        scanbarcode PLcdata(102)
        loadGridData
   End If
   
   If PLcdata(109) = 0 And pulse1 = False Then
      pulse1 = True
      lblGo.Visible = False
      lblNg.Visible = False
   ElseIf PLcdata(109) = 1 And pulse1 = True Then
      pulse1 = False
      lblGo.Visible = True
      GetCounterValue
      txtproductioncounter.Text = Val(txtproductioncounter.Text) + 1
      txtOKCounter.Text = Val(txtOKCounter.Text) + 1
      SaveReport 1
      SaveCounterValue
      loadGridData
   ElseIf PLcdata(109) = 2 And pulse1 = True Then
      pulse1 = False
      GetCounterValue
      lblNg.Visible = True
      txtNGCounter.Text = Val(txtNGCounter.Text) + 1
      SaveReport 2
      SaveCounterValue
      loadGridData
   End If
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "Assign Plc Data"
   Resume Next
End Function
Private Function scanbarcode(count As Integer) As Integer
On Error GoTo Error
Dim result(4) As Integer
txtBarcode(0).BackColor = vbWhite
txtBarcode(1).BackColor = vbWhite
txtBarcode(2).BackColor = vbWhite
txtBarcode(3).BackColor = vbWhite
    Dim i As Integer
    Dim j As Integer
    Dim resultString As String
    
    For j = 0 To count - 1
        Dim str() As String
        resultString = ""
        i = 0
        For i = 0 To 9
            ' Convert integer to 2 bytes
            Dim number As Integer
            number = PLcdata(110 + i + j * 10)
            Dim byte1str As String
            Dim byte2str As String
            Dim byte1 As Byte
            Dim byte2 As Byte
            byte1 = 0
            byte2 = 0
            byte1str = ""
            byte2str = ""
            byte1 = number \ 256 ' Get the first byte
            byte2 = number Mod 256 ' Get the second byte
            If byte1 > 0 Then
            byte1str = Chr(byte1)
            End If
            If byte2 > 0 Then
            byte2str = Chr(byte2)
            
            End If
            resultString = resultString & byte2str & byte1str
        Next i
        str = Split(resultString, ":")
        txtBarcode(j).Text = str(0)
    Next
    j = 0
    For j = 0 To count - 1
        result(j) = CheckBarcodeValidation(txtBarcode(j).Text)
        If result(j) = 1 Then
            txtBarcode(j).BackColor = vbGreen
        Else
            txtBarcode(j).BackColor = vbRed
        End If
    Next
    j = 0
    For j = 0 To count - 1
        If result(j) <> 1 Then
            PLcdata(203) = result(j)
            Exit Function
        End If
    Next
    PLcdata(203) = 1
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "Scan Barcode"
   PLcdata(203) = 2
End Function
Private Function byteBarcode(int1 As Integer) As String


End Function
Private Sub saveToNotepad()
On Error GoTo Error
Dim tempDate, tempCounter As String

tempDate = GetDateCode
MarkDateCode = tempDate
tempCounter = Format(txtproductioncounter.Text, "00000")
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
    Print #iFileNo, tempDate
    Close iFileNo
    Open iFile For Append As iFileNo
    Print #iFileNo, tempCounter
    Close iFileNo
Else
    ErrorLog 0, "Notepad FolderPath not found", "", "", ""
End If
Exit Sub
Error:
  ErrorLog Err.number, Err.Description, Err.Source, "", ""
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
On Error GoTo Error
With VSFModel

.TextMatrix(0, 1) = "Barcode"
.ColWidth(0) = 1000
.Rows = 1
End With
loadGridData
Exit Sub
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "LoadGrid"

End Sub
Private Sub loadGridData()
On Error GoTo Error
With VSFModel
Dim rs As ADODB.Recordset
Dim Sql As String
Dim count As Integer
.Rows = 1
    Sql = "Select * from DummyTable where ModelName='" & ModelName & "' order by id desc"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    count = RecordCount(Sql)
    txtPartScanned = count
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 1) = rs("Barcode")
        .TextMatrix(.Rows - 1, 0) = count - .Rows + 2
        rs.MoveNext
    Loop

End With
Exit Sub
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "loadGridData"
End Sub
Private Function RecordCount(ByVal Sql As String)
On Error GoTo Error
'Dim Sql As String
Dim rs As ADODB.Recordset
Dim Row As Long

'    Sql = "Select * from " & Table
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenStatic, , adCmdText

    Row = Format$(rs.RecordCount)
    rs.Close

    RecordCount = Row

Exit Function
Error:
MsgBox Error, vbInformation
End Function

Private Sub LoadData()
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
Dim strByPass(14) As Integer
Dim j As Integer

    Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    PLcdata(210) = Val(rs("ModelNo"))
    PLcdata(211) = Val(rs("PartInTrey"))
    
Exit Sub
Error:
ErrorLog Err.number, Err.Description, Erl, Me.Name, "LoadData"
Resume Next
End Sub
Private Sub LoadSettingsData()
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
   Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   txtModelDesc.Text = rs("ModelDesc")
   Text1.Text = Val(rs("PartInTrey"))
    If rs("Bypass1") = 1 Then
        SqlBypass = True
    Else
        SqlBypass = False
        If SqlConn = True Then
        Con1.Close
        End If
    End If
    PartNumber = rs("PartNo")
    If rs("Bypass2") = 1 Then
        BarcodeRepeatEnable = True
    Else
        BarcodeRepeatEnable = False
    End If
    If rs("Bypass3") = 1 Then
        BarcodeValidation = True
    Else
        BarcodeValidation = False
    End If
    If rs("Bypass4") = 1 Then
        PartNoValidation = True
    Else
        PartNoValidation = False
    End If
    
Exit Sub
Error:
ErrorLog Err.number, Err.Description, Erl, Me.Name, "LoadSettingsData"
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
Private Sub SaveReport(result As Integer)
'On Error GoTo Error
'Dim Sql As String
'Dim rs As ADODB.Recordset
'   Sql = "Select * from Model_Report"
'   Set rs = New ADODB.Recordset
'   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
'   rs.AddNew
'      rs("ModelName") = ModelName
'      rs("OperatorName") = LoginUser
'      rs("Date") = Format(Now(), "dd/MM/yyyy")
'      rs("Time") = Format(Now(), "hh:mm:ss")
'      rs("Result") = result
'    rs.Update
MoveDummydataToModelReport result
If result = 1 Then
SaveTreyDataInSQl
End If
deletedummyData
End Sub
Private Function SaveToDummy(barcodestr As String, result As Integer)
On Error GoTo Error
Dim Sql As String
Dim rs As ADODB.Recordset
   Sql = "Select * from DummyTable"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   rs.AddNew
      rs("ModelName") = ModelName
      rs("OperatorName") = LoginUser
      rs("Date") = Format(Now(), "dd/MM/yyyy")
      rs("Time") = Format(Now(), "hh:mm:ss")
      rs("Barcode") = barcodestr
      rs("Result") = result
    rs.Update
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "SaveToDummy"
    
End Function


Private Sub SaveCounterValue()
On Error GoTo Error
 Dim ProdDay As String
 SaveSetting App.Title, ModelName, "OkCounter", Val(txtOKCounter.Text)
 SaveSetting App.Title, ModelName, "NGCounter", Val(txtNGCounter.Text)
 SaveSetting App.Title, ModelName, "ProductionCounter", Val(txtproductioncounter.Text)
Exit Sub
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "SaveCounterValue"
End Sub
Private Sub SaveProductioncounter()
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
    Sql = "Select * from Model_Set where ModelName ='" & ModelName & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    rs("productioncounter") = Val(txtproductioncounter.Text)
    rs.Update
    'txtSaveCoupler.Text = Rs("CouplerCounter")
Exit Sub
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "SaveProductioncounter"
    
End Sub
Private Sub GetCounterValue()
On Error Resume Next
Dim ProdDay As String
Dim Today As String
   txtOKCounter.Text = Val(GetSetting(App.Title, ModelName, "OkCounter", 0))
   txtNGCounter.Text = Val(GetSetting(App.Title, ModelName, "NgCounter", 0))
   txtproductioncounter.Text = GetSetting(App.Title, ModelName, "ProductionCounter", 0)
   runningreportdate = Format(Now(), "ddMMyy")
   runningreportshift = getShift
   tempshift = GetSetting(App.Title, ModelName, "saveshift", 0)
   tempDate = GetSetting(App.Title, ModelName, "savedate", 0)
   If tempDate <> runningreportdate Then
      txtOKCounter.Text = 0
      txtNGCounter.Text = 0
      txtproductioncounter.Text = 0
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
Private Function CheckBarcodeValidation(barcodetocheck As String) As Integer
On Error GoTo Error
    CheckBarcodeValidation = 1
    If barcodetocheck = "" Then
        CheckBarcodeValidation = 4
        Exit Function
    End If
    If BarcodeValidation = False Then
        If PartNoValidation = False Then
           If Mid(barcodetocheck, 1, Len(PartNumber)) <> PartNumber Then
            CheckBarcodeValidation = 5
            Exit Function
           End If
        End If
        If BarcodeRepeatEnable = False Then
            If ValidateBarcodeInReport(barcodetocheck) = False Then
                CheckBarcodeValidation = 6
                Exit Function
            ElseIf ValidateBarcodeInDummy(barcodetocheck) = False Then
                CheckBarcodeValidation = 7
                Exit Function
            ElseIf ValidateBarcodeInDummySameModel(barcodetocheck) = False Then
                CheckBarcodeValidation = 8
                Exit Function
            End If
        End If
        If SqlBypass = False Then
            If ValidateBarcodeInSQl(barcodetocheck) = False Then
                CheckBarcodeValidation = 9
                Exit Function
            End If
        End If
    End If
    If CheckBarcodeValidation = 1 Then
        SaveToDummy barcodetocheck, 1
    End If
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "CheckBarcodeValidation"

End Function

Private Function ValidateBarcodeInSQl(barcodetocheck As String) As Boolean
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
SqlConn
If Con1.State = 1 Then
   Sql = "Select * from mesict_AkSensor where sno='" & barcodetocheck & "' and snostatus = 'PASS'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con1, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      ValidateBarcodeInSQl = True
   Else
      ValidateBarcodeInSQl = False
   End If
Else
    ValidateBarcodeInSQl = False
End If
rs.Close
Con1.Close
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "ValidateBarcodeInSQl"
End Function

Private Function SaveTreyDataInSQl() As Boolean
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim Sql As String
SqlConn
If Con1.State = 1 Then
   Sql = "Select * from FCTTesting" ' where sno='" & barcodetocheck & "' and snostatus = 'PASS'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con1, adOpenDynamic, adLockOptimistic
   Sql = "Select * from DummyTable where ModelName ='" & ModelName & "'" ' where sno='" & barcodetocheck & "' and snostatus = 'PASS'"
   Set rs1 = New ADODB.Recordset
   rs1.Open Sql, Con, adOpenDynamic, adLockOptimistic
   Do While rs1.EOF = False
      rs.AddNew
       'rs("status") = "PASS"
       rs("receiveddatetime") = Format(Now, "yyyy-MM-dd hh:mm:ss")
       rs("recdata") = rs1("Barcode")
       'rs("flag") = 0
      rs.Update
    rs1.MoveNext
   Loop
   SaveTreyDataInSQl = True
Else
SaveTreyDataInSQl = False
End If
rs.Close
If Con1.State = 1 Then Con1.Close
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "SaveTreyDataInSQl"
End Function
Private Function ValidateBarcodeInReport(barcodetocheck As String) As Boolean
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
   Sql = "Select * from Model_Report where barcode = '" & barcodetocheck & "' and result =1  and trayOk = '1'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      ValidateBarcodeInReport = False
   Else
      ValidateBarcodeInReport = True
   End If
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "LoadGrid"
End Function
Private Function ValidateBarcodeInDummy(barcodetocheck As String) As Boolean
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
   Sql = "Select * from DummyTable where barcode='" & barcodetocheck & "'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      ValidateBarcodeInDummy = False
   Else
      ValidateBarcodeInDummy = True
   End If
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "ValidateBarcodeInReport"
End Function
Private Function ValidateBarcodeInDummySameModel(barcodetocheck As String) As Boolean
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
   Sql = "Select * from DummyTable where barcode='" & barcodetocheck & "' and ModelName ='" & ModelName & "'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      ValidateBarcodeInDummySameModel = False
   Else
      ValidateBarcodeInDummySameModel = True
   End If
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "ValidateBarcodeInDummySameModel"
End Function
Private Function loaddummydata()
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
Dim Row As Integer
Row = 1
   Sql = "Select * from DummyTable where ModelName ='" & ModelName & "' order by id desc"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
        txtPartScanned.Text = rs.RecordCount
        Do While rs.EOF = False
            VSFModel.Rows = Row + 1
            VSFModel.TextMatrix(Row, 0) = Val(txtPartScanned.Text - Row + 1)
            VSFModel.TextMatrix(Row, 1) = rs("Barcode")
        Loop
   Else
    txtPartScanned.Text = 0
   
   End If
Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "loaddummydata"
End Function
Private Function deletedummyData() As Integer
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
If Con.State = 1 Then
   Sql = "Delete from DummyTable where ModelName ='" & ModelName & "'"
   Con.Execute Sql
   Sql = "Select * from DummyTable where ModelName = '" & ModelName & "'"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If rs.EOF = True Then
    deletedummyData = 1
   Else
    deletedummyData = 2
   End If
Else
    deletedummyData = 3
End If
Exit Function
Error:
  ErrorLog Err.number, Err.Description, Err.Source, Me.Name, "deletedummyData"
  Err.Clear
  deletedummyData = 4
End Function
Private Function MoveDummydataToModelReport(TreyStatus As Integer)
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
Dim rs1 As ADODB.Recordset
Dim Sql1 As String
   
   Sql = "Select * from DummyTable where ModelName ='" & ModelName & "' order by id asc"
   Set rs = New ADODB.Recordset
   rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   Sql1 = "Select * from Model_Report"
   Set rs1 = New ADODB.Recordset
   rs1.Open Sql1, Con, adOpenDynamic, adLockOptimistic
        Do While rs.EOF = False
            rs1.AddNew
            rs1("Date") = rs("Date")
            rs1("Time") = rs("Time")
            rs1("Barcode") = rs("Barcode")
            rs1("Result") = rs("Result")
            rs1("Trayok") = TreyStatus
            rs1("OperatorName") = rs("OperatorName")
            rs1("ModelName") = rs("ModelName")
            rs1.Update
            rs.MoveNext
        Loop

Exit Function
Error:
   ErrorLog Err.number, Err.Description & "---", Erl, Me.Name, "MoveDummydataToModelReport"
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
  ErrorLog Err.number, Err.Description, Err.Source, "", ""
  Err.Clear
 
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
