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
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7035
      ScaleWidth      =   13155
      TabIndex        =   0
      Top             =   120
      Width           =   13215
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   4440
         TabIndex        =   26
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Caption         =   "ByPass"
         Height          =   1695
         Left            =   360
         TabIndex        =   21
         Top             =   3720
         Width           =   6975
         Begin VB.CheckBox chkBypass 
            Caption         =   "Part No Validation Bypass"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   25
            Top             =   1080
            Width           =   2775
         End
         Begin VB.CheckBox chkBypass 
            Caption         =   "Barcode Validation Bypass"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   24
            Top             =   480
            Width           =   2775
         End
         Begin VB.CheckBox chkBypass 
            Caption         =   "Barcode Repeat Enable"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   2775
         End
         Begin VB.CheckBox chkBypass 
            Caption         =   "SQL Validation Bypass"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Width           =   2775
         End
      End
      Begin VB.TextBox txtPartNo 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1800
         TabIndex        =   18
         Top             =   3120
         Width           =   5535
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   19440
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Height          =   1650
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtModelNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3840
            TabIndex        =   15
            Top             =   1200
            Width           =   2865
         End
         Begin VB.TextBox txtModelDesc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1680
            TabIndex        =   12
            Top             =   720
            Width           =   5025
         End
         Begin VB.TextBox txtModelName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1680
            TabIndex        =   11
            Top             =   240
            Width           =   5025
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model No"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   16
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
            TabIndex        =   14
            Top             =   720
            Width           =   1755
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
            TabIndex        =   13
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
         Height          =   6135
         Left            =   7680
         TabIndex        =   7
         Top             =   720
         Width           =   5505
         Begin VSFlex7Ctl.VSFlexGrid VSFModel 
            Height          =   5205
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   5235
            _cx             =   9234
            _cy             =   9181
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
            TabIndex        =   9
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
            Left            =   960
            TabIndex        =   8
            Top             =   5640
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
         Left            =   360
         TabIndex        =   1
         Top             =   5640
         Width           =   7095
         Begin VB.CommandButton CmdClose 
            Caption         =   "&Close"
            Height          =   810
            Left            =   5520
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
            Left            =   360
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
            Left            =   2040
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
            Left            =   360
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
            Left            =   3600
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
      Begin VB.Label lblTorque 
         Caption         =   "Part in 1 Trey"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   27
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblTorque 
         Caption         =   "Part No"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   19
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Setting Screen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         TabIndex        =   17
         Top             =   120
         Width           =   5415
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
Private Sub LoadGrid()
On Error GoTo Error
Exit Sub
Error:
ErrorLog Err.number, Err.Description, Erl, Me.Name, "LoadGrid"
Resume Next
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




Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
 Combo2.Visible = True
 Combo3.Visible = False
 Else
   Combo2.Visible = False
   Combo3.Visible = True
End If
End Sub

Private Sub Combo9_Click()
If Combo9.ListIndex = 0 Then
 Combo8.Visible = True
 Combo7.Visible = False
 Else
   Combo8.Visible = False
   Combo7.Visible = True
End If
End Sub



Private Sub cmdImage_Click()
With CD1
    .DialogTitle = "Select File"
    .Filter = "(*.bmp; *.jpg;)"
    .ShowOpen
    txtImagePath.Text = .FileName
End With
End Sub







'''Private Sub Command4_Click()
''''Dim X, Y As Integer
'''
'''VSFVolt.Rows = ((Val(txtVacFillTime) / Val(txtVacHoldTime))) + 2 '(((Val(txtTestTravel)) * 2) + 1) + 1
'''
'''For i = 1 To VSFVolt.Rows - 1
'''    'VSFVolt.Rows = VSFVolt.Rows + 1
''''    X = ((i * 2) - 1): Y = (i * 2)
'''    VSFVolt.TextMatrix(i, 0) = Format((i - 1) * Val(txtVacHoldTime), "0") 'Format((i - 1) / 2, "0.0") 'i - 1
''''    VSFVolt.TextMatrix(i, 1) = 0 'Format(((X / 100) * 2.45) - 0.2, "0.000")
''''    VSFVolt.TextMatrix(i, 2) = 5 'Format(((Y / 100) * 2.47) + 0.2, "0.000")
''''    VSFVolt.TextMatrix(i, 3) = 0 'Format(((X / 100) * 1.45) - 0.2, "0.000")
''''    VSFVolt.TextMatrix(i, 4) = 5 'Format(((Y / 100) * 1.47) + 0.2, "0.000")
'''Next
'''1
'''
'''End Sub

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
        rs("ModelDesc") = txtModelDesc.Text
        rs("ModelNo") = Val(txtModelNo.Text)
        rs("PartNo") = txtPartNo.Text
        rs("PartInTrey") = Val(Text1.Text)
        For i = 0 To 3
            rs("Bypass" & i + 1) = chkBypass(i).Value
        Next
    rs.Update
    MsgBox UCase("Saved Successfully")
    FillModelGrid
    ResetForm
Exit Sub
Error:
'MsgBox Error, vbInformation
ErrorLog Err.number, Err.Description, Erl, Me.Name, "Save Model Setting"
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
        txtModelDesc.Text = Trim(rs("ModelDesc"))
        txtModelNo.Text = Val(rs("ModelNo"))
        txtPartNo.Text = rs("PartNo")
        
        Text1.Text = Val(rs("PartInTrey"))
        For i = 0 To 3
            chkBypass(i).Value = rs("Bypass" & i + 1)
        Next
           
    Exit Sub
Error:
ErrorLog Err.number, Err.Description, Erl, Me.Name, "LoadData"
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
     Frame4.Visible = False
    End If
End Sub