' Split string by char
Dim str As String
Dim strArr() As String
str = "14/02/2020"
strArr = str.Split("/")
strArr(0) & strArr(1) & strArr(2)

' Check if recordset item is null
rs.Fields.Item("field").Value Is DBNull.Value
