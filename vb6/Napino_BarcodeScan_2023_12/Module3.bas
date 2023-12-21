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
Dim i As Integer
i = 0
ColArrey(i) = "PartNo"
i = i + 1
ColArrey(i) = "ModelNo"
i = i + 1
ColArrey(i) = "PartInTrey"
For j = 0 To 4
i = i + 1
ColArrey(i) = "Bypass" & j + 1
Next
For Row = 0 To i
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next

'----------------Table - Model_Report
TableName = "Model_Report"
ColArrey(1) = "DateCode"
ColArrey(2) = "Barcode"
ColArrey(3) = "Result"
For Row = 1 To 3
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
ColArrey(41) = "ConnString"
For Row = 1 To 41
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
