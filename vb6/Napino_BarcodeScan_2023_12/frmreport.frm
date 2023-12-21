VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport 
   Caption         =   "Report Generation"
   ClientHeight    =   9735
   ClientLeft      =   225
   ClientTop       =   555
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   13260
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10215
      Left            =   120
      ScaleHeight     =   10155
      ScaleWidth      =   19995
      TabIndex        =   0
      Top             =   120
      Width           =   20055
      Begin MSComDlg.CommonDialog cd1 
         Left            =   480
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Left            =   1440
         Top             =   960
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   2415
         Begin VB.Label lblHeading 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   360
            TabIndex        =   4
            Top             =   600
            Width           =   1305
         End
      End
      Begin VB.Frame FrmDateWise 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   2640
         TabIndex        =   1
         ToolTipText     =   "You can access 100 record at a time"
         Top             =   0
         Width           =   11520
         Begin VB.ComboBox CboModelName 
            Height          =   360
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   360
            Width           =   2775
         End
         Begin VB.ComboBox CboReportType 
            Height          =   360
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
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
            Left            =   9360
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1875
         End
         Begin VB.CheckBox ChkDelete 
            Caption         =   "Check To Delete DATA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7560
            TabIndex        =   2
            Top             =   360
            Width           =   1455
         End
         Begin VB.Frame FrameDT 
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
            Height          =   735
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   7215
            Begin MSComCtl2.DTPicker DTFrom 
               Height          =   405
               Left            =   1200
               TabIndex        =   16
               Top             =   240
               Width           =   2085
               _ExtentX        =   3678
               _ExtentY        =   714
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarForeColor=   16711680
               CalendarTitleForeColor=   49344
               Format          =   112525313
               CurrentDate     =   39022
            End
            Begin MSComCtl2.DTPicker DTTo 
               Height          =   405
               Left            =   4440
               TabIndex        =   18
               Top             =   240
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   714
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarForeColor=   16711680
               CalendarTitleForeColor=   49344
               Format          =   112525313
               CurrentDate     =   39022
            End
            Begin VB.Label lblTo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "To"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3600
               TabIndex        =   17
               Top             =   240
               Width           =   240
            End
            Begin VB.Label lblFrom 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "From"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame FrameBC 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   7095
            Begin VB.TextBox txtBarCode 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   22
               Top             =   240
               Width           =   3495
            End
            Begin VB.Label lblBarcode 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Barcode"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   0
               TabIndex        =   23
               Top             =   240
               Width           =   795
            End
         End
         Begin VB.Label lblModel 
            Caption         =   "Model"
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
            Left            =   3600
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Report By"
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
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame FrameB4 
         Height          =   1575
         Left            =   14280
         TabIndex        =   6
         Top             =   0
         Width           =   5625
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   3840
            Picture         =   "frmreport.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   360
            Width           =   1605
         End
         Begin VB.CommandButton cmdExportToExcel 
            Caption         =   "&Export To CSV"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   285
            Picture         =   "frmreport.frx":1144
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   360
            Width           =   1485
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   2040
            Picture         =   "frmreport.frx":2970
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   360
            Width           =   1485
         End
      End
      Begin VB.Frame FrameFL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   19815
         Begin VSFlex7Ctl.VSFlexGrid VSFReport 
            Height          =   8055
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   19575
            _cx             =   34528
            _cy             =   14208
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   16744576
            ForeColorFixed  =   16777215
            BackColorSel    =   16744576
            ForeColorSel    =   -2147483637
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   2
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   500
            Cols            =   10
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   500
            RowHeightMax    =   0
            ColWidthMin     =   1200
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmreport.frx":3AB4
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
            FillStyle       =   1
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   0   'False
            WordWrap        =   -1  'True
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
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Row As Long
Dim Col As Long

Private Sub ExportToCSV()
On Error GoTo Error
Dim Row, Col As Long
Dim strData As String
Dim strLine As String
Dim FilePath As String

    strData = ""
    strData = strData & vbNewLine & vbNewLine
    strData = strData & Format(DTFrom, "dd/mm/yyyy") & " To " & Format(DTTo, "dd/mm/yyyy")
    strData = strData & vbNewLine & vbNewLine

    For Row = 0 To VSFReport.Rows - 1
        strLine = ""
        For Col = 0 To VSFReport.Cols - 1
            strLine = strLine & Trim(VSFReport.TextMatrix(Row, Col)) & ","
        Next
        strData = strData & strLine & vbNewLine
    Next

    With CD1
        .DialogTitle = "Save Report"
        .FileName = ""
        .InitDir = Mid$(App.Path, 1, 3)
        .Filter = "Report Files (.csv)|*.csv"
        .ShowSave
    End With
    If LenB(CD1.FileName) = 0 Then Exit Sub
    FilePath = CD1.FileName

    'Print Report Into File
    Open FilePath$ For Output As #1
        Print #1, strData
    Close #1

Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub CboReportType_Click()

If CboReportType.ListIndex = 3 Then
    FrameDT.Visible = False
    FrameBC.Visible = True

    cbomodelname.Visible = False
    lblModel.Visible = False
    Timer1.Enabled = True
    Timer1.Interval = 1000
Else
    FrameBC.Visible = False
    cbomodelname.Visible = True
    lblModel.Visible = True
    FrameDT.Visible = True


End If

End Sub


Private Sub cmdClose_Click()
    On Error Resume Next
    frmmenu.Show
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Error
Dim Sql As String
Dim SqlWhere(5) As String
    
    If ChkDelete.Value = 0 Then
        MsgBox "Please Check The Box To Delete The Data", vbInformation
        Exit Sub
    End If

VSFReport.Rows = 1
Row = 0

SqlWhere(1) = "Where Date BETWEEN #" & Format(DTFrom.Value, "MM/dd/yyyy") & "# AND #" & Format(DTTo.Value, "MM/dd/yyyy") & "#"
SqlWhere(2) = " and Result = 1"
SqlWhere(3) = " and Result = 2"
SqlWhere(4) = " and ModelName = '" & cbomodelname.Text & "'"
SqlWhere(5) = " where Barcode='" & Trim$(txtBarcode) & "'"

Sql = "Delete from Model_Report "
If CboReportType.ListIndex = 0 Then
    SqlWhere(0) = SqlWhere(1)
    If cbomodelname.ListIndex <> 0 Then SqlWhere(0) = SqlWhere(0) & SqlWhere(4)
    Sql = Sql & SqlWhere(0) '& " order by ID Desc"

ElseIf CboReportType.ListIndex = 1 Then
    SqlWhere(0) = SqlWhere(1) & SqlWhere(2)
    If cbomodelname.ListIndex <> 0 Then SqlWhere(0) = SqlWhere(0) & SqlWhere(4)
    Sql = Sql & SqlWhere(0) '& " order by ID Desc"

ElseIf CboReportType.ListIndex = 2 Then
    SqlWhere(0) = SqlWhere(1) & SqlWhere(3)
    If cbomodelname.ListIndex <> 0 Then SqlWhere(0) = SqlWhere(0) & SqlWhere(4)
    Sql = Sql & SqlWhere(0) '& " order by ID Desc"

ElseIf CboReportType.ListIndex = 3 Then
    SqlWhere(0) = SqlWhere(5) '& SqlWhere(3)
    If cbomodelname.ListIndex <> 0 Then SqlWhere(0) = SqlWhere(0) & SqlWhere(4)
    Sql = Sql & SqlWhere(0) '& " order by ID Desc"
    
End If

Con.Execute Sql

Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub cmdExportToExcel_Click()
On Error Resume Next
    ExportToCSV
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim Row As Double
Dim Sql As String
Dim rs As ADODB.Recordset
Dim TotalRow As Long
Dim SqlWhere(5) As String

VSFReport.Rows = 1
Row = 0

SqlWhere(1) = "Where Date BETWEEN #" & Format(DTFrom.Value, "MM/dd/yyyy") & "# AND #" & Format(DTTo.Value, "MM/dd/yyyy") & "#"
SqlWhere(2) = " and Result = 1"
SqlWhere(3) = " and Result = 2"
SqlWhere(4) = " and ModelName = '" & cbomodelname.Text & "'"
SqlWhere(5) = " where Barcode='" & Trim$(txtBarcode) & "'"
   Sql = "Select * from Model_Report "
   If CboReportType.ListIndex = 0 Then
      SqlWhere(0) = SqlWhere(1)
      If cbomodelname.ListIndex <> 0 Then SqlWhere(0) = SqlWhere(0) & SqlWhere(4)
      Sql = Sql & SqlWhere(0) & " order by ID Desc"

   ElseIf CboReportType.ListIndex = 1 Then
      SqlWhere(0) = SqlWhere(1) & SqlWhere(2)
      If cbomodelname.ListIndex <> 0 Then SqlWhere(0) = SqlWhere(0) & SqlWhere(4)
      Sql = Sql & SqlWhere(0) & " order by ID Desc"

   ElseIf CboReportType.ListIndex = 2 Then
      SqlWhere(0) = SqlWhere(1) & SqlWhere(3)
      If cbomodelname.ListIndex <> 0 Then SqlWhere(0) = SqlWhere(0) & SqlWhere(4)
      Sql = Sql & SqlWhere(0) & " order by ID Desc"

   ElseIf CboReportType.ListIndex = 3 Then
      SqlWhere(0) = SqlWhere(5) '& SqlWhere(3)
      If cbomodelname.ListIndex <> 0 Then SqlWhere(0) = SqlWhere(0) & SqlWhere(4)
      Sql = Sql & SqlWhere(0) & " order by ID Desc"
   End If
    
    TotalRow = RecordCount(Sql)
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    
    Do While rs.EOF = False
      VSFReport.Rows = VSFReport.Rows + 1
      Row = Row + 1
    With VSFReport
      .TextMatrix(Row, 0) = Row
      .TextMatrix(Row, 1) = rs("ModelName")
      .TextMatrix(Row, 2) = rs("Operatorname")
      .TextMatrix(Row, 3) = rs("Date")
      .TextMatrix(Row, 4) = rs("Time")
      If rs("Result") = 1 Then
      .TextMatrix(Row, 5) = "OK"
      ElseIf rs("Result") = 2 Then
      .TextMatrix(Row, 5) = "NG"
      Else
      .TextMatrix(Row, 5) = "-"
      End If
      If rs("trayOk") = "1" Then
      .TextMatrix(Row, 6) = "OK"
      ElseIf rs("trayOk") = "2" Then
      .TextMatrix(Row, 6) = "NG"
      Else
      .TextMatrix(Row, 6) = "-"
      End If
      .TextMatrix(Row, 7) = rs("Barcode")
      
     End With
     If Row > (TotalRow + 1) Then Exit Sub
     rs.MoveNext
    Loop

Exit Sub
Error:
   MsgBox "Error in Searching Record", vbCritical, "Search Error"
End Sub

Private Sub LoadModelCombo(Combo As ComboBox)
Dim Sql As String
Dim rs As ADODB.Recordset
Dim i As Integer

    Combo.Clear
    ' To Fill Combo Box With Switch Names
    Sql = "Select * from Model_Set"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Combo.AddItem "All", 0
    i = 1
    Do While rs.EOF = False
        Combo.AddItem rs("ModelName"), i
        i = i + 1
        rs.MoveNext
    Loop
    Combo.ListIndex = 0
     ' Combo Load End

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

Private Sub Form_Load()
On Error GoTo Error

'Advance
Me.WindowState = 2
Picture1.BorderStyle = 1
Picture1.Appearance = 0
Picture1.BackColor = Me.BackColor '&H80000010
Picture1.Left = (Screen.Width - Picture1.Width) / 2
Picture1.Top = (Screen.Height - Picture1.Height) / 2 - 400
LoadGrid
UserAccess

'VSFReport.Rows = 1s
DTFrom.Format = dtpCustom
DTFrom.CustomFormat = "dd-MM-yyyy"
DTFrom.Value = Date

DTTo.Format = dtpCustom
DTTo.CustomFormat = "dd-MM-yyyy"
DTTo.Value = Date

With CboReportType
    .AddItem "ALL", 0
    .AddItem "OK", 1
    .AddItem "NG", 2
    .AddItem "Barcode", 3
    .ListIndex = 0
End With

LoadModelCombo cbomodelname

Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub LoadGrid()
Dim X As String
With VSFReport
    .Cols = 8
    .Rows = 1
    .FixedRows = 1
    .RowHeightMin = 400
    .RowHeight(0) = 600
    .ExtendLastCol = True
    .HighLight = flexHighlightWithFocus
    .SelectionMode = flexSelectionByRow
    .ScrollBars = flexScrollBarBoth
    
    .TextMatrix(0, 0) = "Sn."
    .TextMatrix(0, 1) = "Model Name"
    .TextMatrix(0, 2) = "Operator"
    .TextMatrix(0, 3) = "Date"
    .TextMatrix(0, 4) = "Time"
    .TextMatrix(0, 5) = "Result"
    .TextMatrix(0, 6) = "Tray Result"
    .TextMatrix(0, 7) = "Barcode"
    
    For Col = 1 To .Cols - 1
        .FixedAlignment(Col) = flexAlignCenterCenter
        .ColAlignment(Col) = flexAlignCenterCenter
        .ColWidth(Col) = 2000
    Next

End With
End Sub

Private Sub UserAccess()

If AccessType = "0" Then 'Disable or Hide For Operator
    cmdDelete.Enabled = False

    
ElseIf AccessType = "1" Then 'Disable or Hide for AccessType 1
    cmdDelete.Enabled = True

ElseIf AccessType = "2" Then 'Show All Which Will Disable or Hide For One
'    CmdCalibration.Visible = True
End If

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    txtBarcode.SetFocus
End Sub
