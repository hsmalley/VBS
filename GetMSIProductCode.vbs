' GetMSIProductCode.vbs

Option Explicit

' Variables
Const msiOpenDatabaseModeReadOnly     = 0

' Get command-line arguements
Dim argCount:argCount = Wscript.Arguments.Count

' Connect to the Windows Installer object.
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError

' Open the database (read-only).
Dim databasePath:databasePath = Wscript.Arguments(0)
Dim openMode : openMode = msiOpenDatabaseModeReadOnly
Dim database : Set database = installer.OpenDatabase(databasePath, openMode) : CheckError

' Extract language info and compose report message
Wscript.Echo "Database (MSI) = "         & databasePath
Wscript.Echo "ProductName    = "         & ProductName(database) 
Wscript.Echo "ProductCode    = "         & ProductCode(database)

' Clean up
Set database = nothing
Wscript.Quit 0

' Get the Property.ProductName value.
Function ProductName(database)
 On Error Resume Next
 Dim view : Set view = database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'ProductName'")
 view.Execute : CheckError
 Dim record : Set record = view.Fetch : CheckError
 If record Is Nothing Then ProductName = "Not specified!" Else ProductName = record.StringData(1)
End Function

' Get the Property.ProductCode value.
Function ProductCode(database)
 On Error Resume Next
 Dim view : Set view = database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'ProductCode'")
 view.Execute : CheckError
 Dim record : Set record = view.Fetch : CheckError
 If record Is Nothing Then ProductCode = "Not specified!" Else ProductCode = record.StringData(1)
End Function

Sub CheckError
 Dim message, errRec
 If Err = 0 Then Exit Sub
 message = Err.Source & " " & Hex(Err) & ": " & Err.Description
 If Not installer Is Nothing Then
  Set errRec = installer.LastErrorRecord
  If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText
 End If
 Fail message
End Sub

Sub Fail(message)
 Wscript.Echo message
 Wscript.Quit 2
End Sub