Attribute VB_Name = "modRegistry"
Option Explicit

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&

Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_SET_VALUE = &H2&
Private Const MAX_PATH = 260&
Private Const REG_DWORD As Long = 4
Private Const REG_SZ = 1
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL

Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Function ReadKey(ByVal KeyName As String, ByVal SubKeyName As String, ByVal ValueName As String, ByVal DefaultValue As String) As String
Dim sBuffer As String
Dim lBufferSize As Long
Dim lphKey As Long
Dim ret As Long
sBuffer = Space(255)
lBufferSize = Len(sBuffer)
ret = RegOpenKey(KeyName, SubKeyName, 0, KEY_READ, lphKey&)
If ret = ERROR_SUCCESS Then
    ret = RegQueryValue(lphKey&, ValueName, 0, REG_SZ, sBuffer, lBufferSize)
    ret = RegCloseKey(lphKey&)
    Else
    ret = RegCloseKey(lphKey&)
    End If
sBuffer = Trim(sBuffer)
If sBuffer <> "" Then
    sBuffer = Left(sBuffer, Len(sBuffer) - 1)
    Else
    sBuffer = DefaultValue
    End If
ReadKey = sBuffer
End Function

Public Function WriteKey(ByVal KeyName As String, ByVal SubKeyName As String, ByVal ValueName As String, ByVal KeyValue As String) As Long
Dim lphKey As Long
Dim ret As Long
ret = RegCreateKey&(KeyName, SubKeyName, lphKey&)
If ret = ERROR_SUCCESS Then
    ret = RegSetValue&(lphKey&, ValueName, REG_SZ, KeyValue, 0&)
    Else
    ret = RegCloseKey(lphKey&)
    End If
WriteKey = ret
End Function

Public Function DeleteKey(ByVal KeyName As String, ByVal SubKeyName As String) As Long
Dim lphKey As Long
Dim ret As Long
ret = RegOpenKey(KeyName, SubKeyName, 0, KEY_WRITE, lphKey&)
If ret = ERROR_SUCCESS Then
    ret = RegDeleteKey(lphKey&, "") 'delete the key
    ret = RegCloseKey(lphKey&)
    End If
DeleteKey = ret
End Function

Public Sub MakeFileAssociation(Extension As String, PathToApplication As String, ApplicationName As String, Description As String, Optional FullIconPath As String)
Dim ret As Long
Dim sKeyName As String
Dim sKeyValue As String
If Left(PathToApplication, 1) <> "\" Then PathToApplication = PathToApplication & "\"
'Create a Root entry called .XXX associated with application name
sKeyName = "." & Extension
sKeyValue = ApplicationName
ret = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
'Set application key and file description
sKeyName = ApplicationName
sKeyValue = Description
ret = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
'This sets the default icon for XXX_auto_file
If FullIconPath <> "" Then
    sKeyName = ApplicationName & "\DefaultIcon"
    sKeyValue = FullIconPath & ",0"
    ret = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
    End If
'This sets the command line for XXX_auto_file
sKeyName = ApplicationName & "\shell\open\command"
sKeyValue = Chr(34) & PathToApplication & ApplicationName & ".exe" & Chr(34) & " %1"
ret = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
End Sub

Public Sub DeleteFileAssociation(Extension As String)
Dim Application As String
Dim ret As Long
'check if filetype is registred
Application = ReadKey(HKEY_CLASSES_ROOT, "." & Extension, "", "")
If Application <> "" Then
    'delete file extension
    ret = DeleteKey(HKEY_CLASSES_ROOT, "." & Extension)
    'delete command lines
    ret = DeleteKey(HKEY_CLASSES_ROOT, Application)
    End If
End Sub
