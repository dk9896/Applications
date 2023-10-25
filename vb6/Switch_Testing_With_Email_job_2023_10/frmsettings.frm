VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmsettings 
   Caption         =   "Setting Test Parameters"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   13260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmsettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   13260
   Begin VB.PictureBox Picture1 
      Height          =   10575
      Left            =   120
      ScaleHeight     =   10515
      ScaleWidth      =   19875
      TabIndex        =   0
      Top             =   120
      Width           =   19935
      Begin VB.OptionButton Opt5V 
         Caption         =   "5 V"
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
         Left            =   6360
         TabIndex        =   67
         Top             =   6960
         Width           =   1095
      End
      Begin VB.OptionButton Opt12V 
         Caption         =   "12 V"
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
         Left            =   4320
         TabIndex        =   66
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox txtMVDMax 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6360
         TabIndex        =   64
         Text            =   "000.0"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox txtCurrentMax 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6360
         TabIndex        =   63
         Text            =   "000.0"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox txtPositionMax 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   6360
         TabIndex        =   62
         Text            =   "0.000"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtPositionMax 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   6360
         TabIndex        =   61
         Text            =   "000"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox txtMVDMin 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4320
         TabIndex        =   59
         Text            =   "000.0"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox txtCurrentMin 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4320
         TabIndex        =   57
         Text            =   "000.0"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox txtTestCycle 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   55
         Text            =   "000.0"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtTestSpeed 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7920
         TabIndex        =   53
         Text            =   "000.0"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtHomeSpeed 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   51
         Text            =   "000.0"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtMaxPos 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7920
         TabIndex        =   49
         Text            =   "000.00"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtHomePos 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   45
         Text            =   "000.0"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtPositionMin 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   4320
         TabIndex        =   44
         Text            =   "0.000"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtPositionMin 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   4320
         TabIndex        =   43
         Text            =   "000"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Frame Frame14 
         Caption         =   "Printer Detail"
         ForeColor       =   &H000040C0&
         Height          =   2055
         Left            =   14760
         TabIndex        =   28
         Top             =   2160
         Width           =   5055
         Begin VB.TextBox txtVandorId 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   40
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox txtPartNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   31
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtSerialNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   30
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtHardwareVersion 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   29
            Top             =   1200
            Width           =   2895
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Code"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   102
            Left            =   120
            TabIndex        =   39
            Top             =   1680
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cust Part No"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   79
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial Starting Text"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   77
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   1665
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index AR"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   75
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   795
         End
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   19440
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame11 
         Caption         =   "Bypasses"
         ForeColor       =   &H000040C0&
         Height          =   1335
         Left            =   360
         TabIndex        =   18
         Top             =   7920
         Width           =   6855
         Begin VB.CheckBox chkbypass 
            Caption         =   "Cycle Time"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   11
            Left            =   5040
            TabIndex        =   42
            Top             =   960
            Width           =   1455
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Presure Switch"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   10
            Left            =   5040
            TabIndex        =   41
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Cavity - 4"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   9
            Left            =   5040
            TabIndex        =   38
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Cavity - 3"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   8
            Left            =   3600
            TabIndex        =   27
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Cavity - 1"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   6
            Left            =   3600
            TabIndex        =   26
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Cavity - 2"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   7
            Left            =   3600
            TabIndex        =   25
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Scanner/ Printer"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   5
            Left            =   1680
            TabIndex        =   24
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "PID - 4"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   4
            Left            =   1680
            TabIndex        =   23
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "PID -3"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   3
            Left            =   1680
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "PID -2"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "PID - 1"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Limit Switch"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2130
         Left            =   14760
         TabIndex        =   11
         Top             =   0
         Width           =   5055
         Begin VB.CommandButton cmdImage 
            Caption         =   "...."
            Height          =   240
            Left            =   4560
            TabIndex        =   37
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txtImagePath 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            TabIndex        =   35
            Top             =   1680
            Width           =   2985
         End
         Begin VB.TextBox txtModelNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3600
            TabIndex        =   16
            Top             =   1200
            Width           =   1305
         End
         Begin VB.TextBox txtModelDesc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            TabIndex        =   13
            Top             =   720
            Width           =   3465
         End
         Begin VB.TextBox txtModelName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   3465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Image Path"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   1680
            Width           =   1875
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model No"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   1875
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model Desc"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1875
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model Name"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Existing Models"
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
         Height          =   5175
         Left            =   14760
         TabIndex        =   7
         Top             =   4200
         Width           =   5025
         Begin VSFlex7Ctl.VSFlexGrid VSFModel 
            Height          =   4365
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   4755
            _cx             =   8387
            _cy             =   7699
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
            FormatString    =   $"frmsettings.frx":116A
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
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To Edit Model Double Click or Press Enter on Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   465
            Left            =   480
            TabIndex        =   10
            Top             =   6720
            Width           =   3705
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Double Click on the Row to get details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   9
            Left            =   600
            TabIndex        =   9
            Top             =   4800
            Width           =   3915
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame5 
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
         Left            =   14760
         TabIndex        =   1
         Top             =   9240
         Width           =   5055
         Begin VB.CommandButton CmdClose 
            Caption         =   "&Close"
            Height          =   810
            Left            =   3720
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmsettings.frx":11D9
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Close Screen"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "&Reset"
            Height          =   810
            Left            =   120
            MaskColor       =   &H00404040&
            Picture         =   "frmsettings.frx":1E1B
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Reset All"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   810
            Left            =   1320
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmsettings.frx":317D
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton cmdAddRow 
            Caption         =   "&Add Row"
            Height          =   810
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmsettings.frx":3DBF
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Add new Line"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmdDeleteRow 
            Caption         =   "&Delete Row"
            Height          =   810
            Left            =   2400
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmsettings.frx":4A01
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Delete Record"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Label Label12 
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   69
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   68
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Test Voltage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   65
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "MVD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   60
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   58
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Test Cycle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   56
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Test Speed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   54
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Home Speed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   52
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Max Pos."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   50
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Home Pos."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   48
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Contact Change Pos. (fwd)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   47
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label Label16 
         Caption         =   "Contact Change Pos. (Rvs)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   46
         Top             =   4440
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Row As Long
Dim Col As Long

Private Sub CboSensorType_Click()

Select Case CboSensorType.ListIndex
    Case 0
        VSFChannel.Cell(flexcpBackColor, 3, 3, 4, 4) = vbWhite
        VSFChannel.Cell(flexcpBackColor, 6, 3, 6, 4) = vbWhite
'        VSFChannel.Cell(flexcpBackColor, 10, 3, 10, 4) = &H404040
    Case 1
        VSFChannel.Cell(flexcpBackColor, 3, 3, 4, 4) = &H404040
        VSFChannel.Cell(flexcpBackColor, 6, 3, 6, 4) = &H404040
'        VSFChannel.Cell(flexcpBackColor, 10, 3, 10, 4) = vbWhite
End Select

End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub DeleteCSV(ByVal FileName As String)
Dim FSO As New FileSystemObject
Dim FilePath As String
    
    FilePath = App.Path & "\ExCelMaster\" & FileName & ".csv"
    
    If FSO.FileExists(FilePath) = True Then
        FSO.DeleteFile FilePath, True
    End If

End Sub

Private Sub WriteCSV(ByVal Grid As VSFlexGrid, ByVal FileName As String)
On Error GoTo Error
Dim Row, Col As Long
Dim strData As String
Dim strLine As String
Dim FilePath As String
    
    FilePath = App.Path & "\ExCelMaster\" & FileName & ".csv"
    
    For Row = 0 To Grid.Rows - 1
        strLine = ""
        For Col = 0 To Grid.Cols - 1
            If Col <> 0 Then strLine = strLine & ","
            strLine = strLine & Trim(Grid.TextMatrix(Row, Col))
        Next
        strData = strData & strLine & vbNewLine
    Next
    
    'Print Report Into File
    Open FilePath$ For Output As #1
        Print #1, strData
    Close #1

Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub ReadCSV(ByVal Grid As VSFlexGrid, ByVal FileName As String)
On Error Resume Next
Dim iFile As Integer
Dim Row, Col As Long
Dim strData As String
Dim strLine() As String
Dim strArray() As String
Dim FilePath As String

    FilePath = App.Path & "\ExCelMaster\" & FileName & ".csv"

    'Read the entire file
    iFile = FreeFile
    Open FilePath For Input As #iFile
        strData = Input(LOF(iFile), iFile)
    Close iFile
    'Split the results into separate lines
    strLine = Split(strData, vbCrLf)
    
    For Row = 0 To UBound(strLine)
        strArray = Split(strLine(Row), ",")
        For Col = 0 To UBound(strArray)
            Grid.TextMatrix(Row, Col) = strArray(Col)
        Next
    Next

ErrorHandler:
Close iFile
End Sub
Private Sub cmdImage_Click()
With CD1
    .DialogTitle = "Select File"
    .Filter = "(*.bmp; *.jpg;)"
    .ShowOpen
    txtImagePath.Text = .FileName
End With
End Sub
Private Sub VSFModel_DblClick()
Dim Row As Integer

Row = VSFModel.Row
txtModelName = Trim(VSFModel.TextMatrix(Row, 1))

If Row >= 1 Then LoadData
    
End Sub

Private Sub FillModelGrid()
Dim Sql As String
Dim rs As ADODB.Recordset
Dim Row As Integer
    
    VSFModel.Rows = 1
    
    Sql = "Select * from Model_Set order by ModelName"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    
    Do While rs.EOF = False
        VSFModel.Rows = VSFModel.Rows + 1
        Row = VSFModel.Rows - 1
        VSFModel.TextMatrix(Row, 0) = Trim(Row)
        VSFModel.TextMatrix(Row, 1) = Trim(rs("ModelName"))
        rs.MoveNext
    Loop
    
End Sub

Private Sub cmdAddRow_Click()

    VSFModel.Rows = VSFModel.Rows + 1
    VSFModel.Select VSFModel.Rows - 1, 1
    VSFModel.TopRow = VSFModel.Rows - 1
    VSFModel.Cell(flexcpBackColor, VSFModel.Rows - 1, 1, VSFModel.Rows - 1, VSFModel.Cols - 1) = RGB(220, 220, 220)
    VSFModel.LeftCol = 0
    VSFModel.SetFocus
    VSFModel.TextMatrix(VSFModel.Rows - 1, 0) = Trim(VSFModel.Rows - 1)
    VSFModel.TextMatrix(VSFModel.Rows - 1, 1) = "Fill The Required Fields"
    ResetForm
    
End Sub

Private Sub cmdDeleteRow_Click()
Dim Sql As String
Dim rs As ADODB.Recordset
   
    If Trim(txtModelDesc) = "" Then
        MsgBox "No Model Is Selected"
    End If
  
    If MsgBox(UCase("Do You Want To Delete?"), vbYesNo + vbInformation) = vbYes Then
  
        Sql = "Select * from Model_Set where ModelName='" & Trim(txtModelName) & "'"
        Set rs = New ADODB.Recordset
        rs.Open Sql, Con, adOpenForwardOnly, adLockOptimistic
        If rs.EOF = True Then Exit Sub
        rs.Delete
        rs.Update
        
        DeleteCSV Trim$(txtModelName) & "-FORCE"
        DeleteCSV Trim$(txtModelName) & "-TRAVEL"
    End If


    ResetForm
    FillModelGrid

End Sub

Private Sub cmdReset_Click()
    If MsgBox(UCase("Reset the form?"), vbYesNo) = vbYes Then
       FillModelGrid
       ResetForm
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmmenu.Show
End Sub

Private Sub CmdSave_Click()
On Error GoTo Error
Dim Sql As String
Dim rs As ADODB.Recordset
Dim O, P As String
    If CheckValidEntry = False Then Exit Sub
    
    Sql = "Select * from Model_Set where ModelName = '" & Trim(txtModelName.Text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    If rs.EOF = True Then
        MsgBox "Creating New Record", vbOKOnly
        rs.AddNew
    ElseIf rs.EOF = False Then
         MsgBox "Record with this Model Name Exist, Updating the record", vbOKOnly
    End If
    rs("ModelName") = Trim(txtModelName.Text)
    rs("ModelDesc") = Trim(txtModelDesc.Text)
    rs("ServoHomePos") = Format(Val(txtHomePos.Text), "0")
rs("ServoMaxPos") = Format(Val(txtMaxPos.Text), "0")
rs("ServoHomeSpeed") = Format(Val(txtHomeSpeed.Text), "0")
rs("ServoTestSpeed") = Format(Val(txtTestSpeed.Text), "0")
rs("TestCycle") = Format(Val(txtTestCycle.Text), "0")
rs("FwdChangeoverMin") = Format(Val(txtPositionMin(0).Text), "0")
rs("FwdChangeoverMax") = Format(Val(txtPositionMax(0).Text), "0")
rs("RvsChangeoverMin") = Format(Val(txtPositionMin(0).Text), "0")
rs("RvsChangeoverMax") = Format(Val(txtPositionMax(1).Text), "0")
rs("CurrentMin") = Format(Val(txtCurrentMin.Text), "0")
rs("CurrentMax") = Format(Val(txtCurrentMax.Text), "0")
rs("MvdMin") = Format(Val(txtMVDMin.Text), "0")
rs("MvdMax") = Format(Val(txtMVDMax.Text), "0")
If Opt12V.Value = True Then
    rs("Testvoltage") = 1
ElseIf Opt5V.Value = True Then
    rs("Testvoltage") = 2
End If

'    rs("PrintPartNo") = txtPartNo.Text
'    rs("HardwareNo") = txtHardwareVersion.Text
'    rs("SerialStartingtxt") = txtSerialNo.Text
'    rs("VendorId") = txtVandorId.Text
'    rs("ModelNo") = txtModelNo.Text
'    rs("PartImage") = txtImagePath.Text
'    rs("PrinterBypass") = Val(chkbypass(5).Value)
'    For i = 0 To 11
'     rs("Bypass" & i + 1) = Val(chkbypass(i).Value)
'    Next
    
    rs.Update
'    WriteCSV VSFData1, Trim$(txtModelName)
    MsgBox UCase("Saved Successfully")
    FillModelGrid
    ResetForm
Exit Sub
Error:
'MsgBox Error, vbInformation
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "Save Model Setting"
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo Error

'Settings
Me.WindowState = 2
Me.BackColor = &H80000010
Picture1.BorderStyle = 1
Picture1.Appearance = 0
Picture1.BackColor = vbButtonFace
Picture1.Left = (Screen.Width - Picture1.Width) / 2
Picture1.Top = (Screen.Height - Picture1.Height) / 2 - 400
FillModelGrid
LoadGrid
UserAccess

Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub LoadData()
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
    
    Sql = "Select * from Model_Set where ModelName ='" & Trim(txtModelName.Text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    
    'txtModelName.Text = Trim(Rs("ModelName"))
    txtModelDesc.Text = Trim(rs("ModelDesc"))
    'txtPartNo.Text = Rs("PrintPartNo")
    For i = 0 To 3
    With VSCavity(i)
        Col = 4
        .TextMatrix(2, Col) = rs("TestVolOnMin" & i + 1)
        .TextMatrix(3, Col) = rs("TestCurOnMin" & i + 1)
        .TextMatrix(4, Col) = rs("MVDOnMin" & i + 1)
        .TextMatrix(5, Col) = rs("ContBrkPntOnMin" & i + 1)
        .TextMatrix(6, Col) = rs("OpeForceOnMin" & i + 1)
        '.TextMatrix(7, Col) = rs("ContResOnMin" & i + 1)
        Col = 5
        .TextMatrix(2, Col) = rs("TestVolOnMax" & i + 1)
        .TextMatrix(3, Col) = rs("TestCurOnMax" & i + 1)
        .TextMatrix(4, Col) = rs("MVDOnMax" & i + 1)
        .TextMatrix(5, Col) = rs("ContBrkPntOnMax" & i + 1)
        .TextMatrix(6, Col) = rs("OpeForceOnMax" & i + 1)
        '.TextMatrix(7, Col) = rs("ContResOnMax" & i + 1)
        Col = 6
        .TextMatrix(2, Col) = rs("TestVolOnTarget" & i + 1)
        .TextMatrix(3, Col) = rs("TestCurOnTarget" & i + 1)
        .TextMatrix(4, Col) = rs("MVDOnTarget" & i + 1)
        .TextMatrix(5, Col) = rs("ContBrkPntOnTarget" & i + 1)
        .TextMatrix(6, Col) = rs("OpeForceOnTarget" & i + 1)
        '.TextMatrix(7, Col) = rs("ContResOnTarget" & i + 1)
        Col = 1
        .TextMatrix(2, Col) = rs("TestVolOffMin" & i + 1)
        .TextMatrix(3, Col) = rs("TestCurOffMin" & i + 1)
        .TextMatrix(4, Col) = rs("MVDOffMin" & i + 1)
        .TextMatrix(5, Col) = rs("ContBrkPntOffMin" & i + 1)
        .TextMatrix(6, Col) = rs("OpeForceOffMin" & i + 1)
        '.TextMatrix(7, Col) = rs("ContResOffMin" & i + 1)
        Col = 2
        .TextMatrix(2, Col) = rs("TestVolOffMax" & i + 1)
        .TextMatrix(3, Col) = rs("TestCurOffMax" & i + 1)
        .TextMatrix(4, Col) = rs("MVDOffMax" & i + 1)
        .TextMatrix(5, Col) = rs("ContBrkPntOffMax" & i + 1)
        .TextMatrix(6, Col) = rs("OpeForceOffMax" & i + 1)
        '.TextMatrix(7, Col) = rs("ContResOffMax" & i + 1)
        Col = 3
        .TextMatrix(2, Col) = rs("TestVolOffTarget" & i + 1)
        .TextMatrix(3, Col) = rs("TestCurOffTarget" & i + 1)
        .TextMatrix(4, Col) = rs("MVDOffTarget" & i + 1)
        .TextMatrix(5, Col) = rs("ContBrkPntOffTarget" & i + 1)
        .TextMatrix(6, Col) = rs("OpeForceOffTarget" & i + 1)
        '.TextMatrix(7, Col) = rs("ContResOffTarget" & i + 1)
     End With
    Next

    txtHomingPos.Text = Format(rs("ServoHomePos"), "0.000")
    txtHomeSpeed.Text = Format(rs("ServoHomeSpeed"), "0.000")
    txtTestPos.Text = Format(rs("ServoTestPos"), "0.000")
    txtTestingSpeed.Text = Format(rs("ServoTestspeed"), "0.000")
    txtFastPos.Text = Format(rs("ServoFastPos"), "0.000")
    txtFastSpeed.Text = Format(rs("ServoFastSpeed"), "0.000")
    'txtMaxtravel.Text = Format(rs("ServoMaxTravel"), "0.000")
    
    txtModelNo.Text = rs("ModelNo")
    'txtVandorId.Text = Rs("VendorId")
    'Rs("BatchCounter").Text
    'Rs("CouplerCounter") = .Text
    txtImagePath.Text = rs("PartImage")
    'Rs("productioncounter") =
    For i = 0 To 11
      chkbypass(i).Value = Val(rs("Bypass" & i + 1))
    Next
    'chkbypass(1).Value = Val(Rs("LSBypass"))
    'chkbypass(2).Value = Val(Rs("WLCBypass"))
    'chkbypass(3).Value = Val(Rs("BSBypass"))
    'chkbypass(4).Value = Val(Rs("PrinterBypass"))
    'chkbypass(5).Value = Val(Rs("ICBypass"))
    'chkbypass(6).Value = Val(Rs("ScannerBypass"))
    'chkbypass(7).Value = Val(Rs("PIDByPass"))
    'chkbypass(8).Value = Val(Rs("PressureGuageByPass"))
    'chkbypass(9).Value = Val(Rs("UpperCoverByPass"))
       
    Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
Resume Next
End Sub

Private Function CheckValidEntry() As Boolean
    
    If ValidLen(3, 30, txtModelName) = False Then Exit Function
    If ValidLen(1, 40, txtModelDesc) = False Then Exit Function
    'If ValidLen(4, 4, txtvendorCode) = False Then Exit Function
    'If ValidLen(1, 1, txtlinecode) = False Then Exit Function
    'If ValidLen(11, 11, txtPartNo) = False Then Exit Function
    'If ValidLen(5, 5, txtLastPartno) = False Then Exit Function
    
'    If ValidEntry(0, 320, txtDataMin3) = False Then Exit Function
'    If ValidEntry(0, 320, txtDataMax3) = False Then Exit Function
'
'    If ValidLen(10, 10, txtDataMin4) = False Then Exit Function
'    If ValidLen(8, 8, txtDataMax4) = False Then Exit Function
'
'
'
'    If ValidEntry(0, 180, txtServoFastSpeed) = False Then Exit Function
'    If ValidEntry(0, 90, txtServoFastDegree) = False Then Exit Function
'    If ValidEntry(0, 90, txtServoSlowSpeed) = False Then Exit Function
'    If ValidEntry(0, 320, txtClampingTime) = False Then Exit Function
'
'    If ValidEntry(1, 90, txtTestCycle) = False Then Exit Function
'    If ValidEntry(0, 30000, txtCameraJob) = False Then Exit Function
'    If ValidEntryGrd(VSFChannel, 1, 1, 0, 300) = False Then Exit Function
'    If ValidEntryGrd(VSFChannel, 1, 2, 0, 300) = False Then Exit Function
'    If ValidEntryGrd(VSFChannel, 1, 3, 0, 300) = False Then Exit Function
'    If ValidEntryGrd(VSFChannel, 1, 4, 0, 300) = False Then Exit Function

   
CheckValidEntry = True
End Function

Private Function ValidEntryGrd(Grid As VSFlexGrid, Row, Col As Integer, Min, Max As String) As Boolean

    If IsNumeric(Grid.TextMatrix(Row, Col)) = False Or _
        Val(Grid.TextMatrix(Row, Col)) < Val(Min) Or _
        Val(Grid.TextMatrix(Row, Col)) > Val(Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max), vbCritical
        Grid.Select Row, Col
        Grid.EditCell
        Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
        ValidEntryGrd = False
    Else
        Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbWhite
        ValidEntryGrd = True
    End If

End Function

Private Function ValidEntry(Min, Max As Double, Text As TextBox) As Boolean

    If IsNumeric(Text) = False Or (Val(Text) < Min Or Val(Text) > Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max), vbInformation
        Text.SetFocus
        Text.BackColor = vbRed
        ValidEntry = False
    Else
        Text.BackColor = vbWhite
        ValidEntry = True
    End If

End Function

Private Function ValidLen(Min, Max As Long, Text As TextBox) As Boolean

    If Trim(Text) = "" Or (Len(Text) < Min Or Len(Text) > Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max & " Characters"), vbCritical
        Text.SetFocus
        Text.BackColor = vbRed
        ValidLen = False
    Else
        Text.BackColor = vbWhite
        ValidLen = True
    End If

End Function

Private Sub ResetForm()
Dim txt As Control

For Each txt In Me
    If TypeOf txt Is TextBox Then
        txt.Text = ""
    End If

    If TypeOf txt Is CheckBox Then
        txt.Value = 0
    End If

    If TypeOf txt Is ComboBox Then
        txt.ListIndex = 0
    End If
Next



'LoadGrid

End Sub
Public Sub UserAccess()
    If AccessType < 2 Then
    For i = 0 To 3
'        lblcurentoffset(i).Visible = False
'        lblvoltageoffset(i).Visible = False
'        txtCurrentOffset(i).Visible = False
'        txtVoltageOffset(i).Visible = False
        VSCavity(i).ColHidden(3) = True
        VSCavity(i).ColHidden(6) = True
    Next
    Frame11.Visible = False
    End If
End Sub
