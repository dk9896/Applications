Attribute VB_Name = "Module3"
Public Con100 As ADODB.Connection

Public Sub Conn_2007()
On Error GoTo Error
Dim StrMdbPath, StrConn As String

    StrMdbPath = App.Path & "\Database\" & App.Title & "_DB.mdb"
    StrConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & StrMdbPath & ";Jet OLEDB:Database Password=authentic;"
    Set Con100 = New ADODB.Connection
    Con100.Open StrConn

Exit Sub
Error:
MsgBox "Office 2007 Error", vbInformation
End Sub

Public Function CreateField(ByVal DBConn, strTable, strField, strType As String) As Boolean
Dim Sql As String
On Error GoTo error1
    If (strField <> vbNullString) Then
        Sql = "ALTER TABLE " & strTable & " ADD COLUMN " & strField & " " & strType
        DBConn.Execute Sql
        Sql = "UPDATE " & strTable & " Set " & strField & " = 0"
        DBConn.Execute Sql
        CreateField = True
    End If
Exit Function
error1:
MsgBox (Err.Description)

End Function

Public Function FieldExists(ByVal DBConn, TableName, FieldName As String) As Boolean
Dim rs As New ADODB.Recordset
Dim FLD As ADODB.Field

rs.Open TableName, DBConn, adOpenStatic, adLockReadOnly, adCmdTable
For Each FLD In rs.Fields
    If LCase(FLD.Name) = LCase(FieldName) Then
        FieldExists = True
        Exit For
    End If
Next

End Function


Public Sub Make_Column()
On Error GoTo Error
Dim Sql As String
Dim TableName As String
Dim ColName As String
Dim Row As Long
Dim Col As Long
Dim ColArrey(200) As String

'----------------Table - Model_Set
TableName = "Model_Set"
'  "Printtype"
'    Rs ("IDNo")
'    Rs("LastPartNo") = txtLastPartno.Text
'    Rs("PartNo") = txtPartNo.Text
'    Rs("Darkness") = cboDarkness.ListIndex
'    Rs("RejectionBypass") = Check1.Value
'    Rs("Vendorcode") = txtvendorCode.Text
'    Rs("linecode") = txtlinecode.Text
ColArrey(1) = "ServoHomePos"
ColArrey(2) = "ServoMaxPos"
ColArrey(3) = "ServoHomeSpeed"
ColArrey(4) = "ServoTestSpeed"
ColArrey(5) = "TestCycle"
ColArrey(6) = "FwdChangeoverMin"
ColArrey(7) = "FwdChangeoverMax"
ColArrey(8) = "RvsChangeoverMin"
ColArrey(9) = "RvsChangeoverMax"
ColArrey(10) = "CurrentMin"
ColArrey(11) = "CurrentMax"
ColArrey(12) = "MvdMin"
ColArrey(13) = "MvdMax"
ColArrey(14) = "Testvoltage"
ColArrey(15) = "VendorId"
ColArrey(16) = "PrintSwitchName"
ColArrey(17) = "PrintLineCode"
ColArrey(18) = "CouplerCounter"
ColArrey(18) = "BatchCounter"


For i = 0 To 10
ColArrey(19 + i) = "Bypass" & i + 1
Next
For Row = 1 To 29
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next

'----------------Table - Model_Report
TableName = "Model_Report"
      ColArrey(1) = "TestVoltage"
      ColArrey(2) = "TestVoltageResult"
      ColArrey(3) = "FwdPosition"
      ColArrey(4) = "FwdPositionResult"
      ColArrey(5) = "FwdCurrent"
      ColArrey(6) = "FwdCurrentResult"
      ColArrey(7) = "FwdMVD"
      ColArrey(8) = "FwdMVDResult"
      ColArrey(9) = "RwdPosition"
      ColArrey(10) = "RwdPositionResult"
      ColArrey(11) = "RwdCurrent"
      ColArrey(12) = "RwdCurrentResult"
      ColArrey(13) = "RwdMVD"
      ColArrey(14) = "RwdMVDResult"
      
For Row = 1 To 14
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next


'----------------Table - Model_Set
TableName = "Common_Set"
ColArrey(1) = "WebApiLink"
ColArrey(2) = "SenderEmail"
ColArrey(3) = "SenderPassword"
ColArrey(4) = "ToEmail1"
ColArrey(5) = "ToEmail2"
ColArrey(6) = "ToEmail3"
ColArrey(7) = "ToEmail4"
ColArrey(8) = "ToEmail5"
ColArrey(9) = "ToEmail6"
ColArrey(10) = "ToEmail7"
ColArrey(11) = "EmailBypass"
ColArrey(12) = "EmailBypass1"
ColArrey(13) = "EmailBypass2"
ColArrey(14) = "EmailBypass3"
ColArrey(15) = "EmailBypass4"
ColArrey(16) = "EmailBypass5"
ColArrey(17) = "EmailBypass6"
ColArrey(18) = "EmailBypass7"
ColArrey(19) = "cycletime"
ColArrey(20) = "Break1Enable"
ColArrey(21) = "Break2Enable"
ColArrey(22) = "Break3Enable"
ColArrey(23) = "Break4Enable"
ColArrey(24) = "Break5Enable"
ColArrey(25) = "Break1Start"
ColArrey(26) = "Break1End"
ColArrey(27) = "Break2Start"
ColArrey(28) = "Break2End"
ColArrey(29) = "Break3Start"
ColArrey(30) = "Break3End"
ColArrey(31) = "Break4Start"
ColArrey(32) = "Break4End"
ColArrey(33) = "Break5Start"
ColArrey(34) = "Break5End"
ColArrey(35) = "Shift1Start"
ColArrey(36) = "Shift1End"
ColArrey(37) = "Shift2Start"
ColArrey(38) = "Shift2End"
ColArrey(39) = "Shift3Start"
ColArrey(40) = "Shift3End"

For Row = 1 To 40
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next
    


'------------------------------------



'TableName = "user_list"
'ColName = "AccessType"
'If FieldExists(Con100, TableName, ColName) = False Then
'    X = CreateField(Con100, TableName, ColName, "varchar(255) DEFAULT 0")
'End If
'-=========================================

'Sql = "create table Common_Set (ID Counter)"
'Con100.Execute Sql

'TableName = "Common_Set"
'ColName = "SetType"
'If FieldExists(Con100, TableName, ColName) = False Then
'    X = CreateField(Con100, TableName, ColName, "varchar(255) DEFAULT 0")
'End If
'
''Sql = "Update Common_Set Set SetType='CommonSet'"
''Con100.Execute Sql
'
'ColName = "ComPort1"
'If FieldExists(Con100, TableName, ColName) = False Then
'    X = CreateField(Con100, TableName, ColName, "varchar(255) DEFAULT 0")
'End If

'Con100.Close

Exit Sub
Error:
MsgBox ("Error")
'Con100.Close
End Sub
