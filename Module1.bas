Attribute VB_Name = "Module1"
Option Explicit


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1 'Unicode nul terminated string
Public Const REG_BINARY = 3 'Free form binary
Public Const REG_DWORD = 4 '32-bit number
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal HKEY As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
    lpData As Any, lpcbData As Long) As Long

Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
(ByVal HKEY As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long


Public Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal HKEY As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" _
Alias "RegCreateKeyA" (ByVal HKEY As Long, ByVal lpSubKey _
As String, phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" _
Alias "RegDeleteKeyA" (ByVal HKEY As Long, ByVal lpSubKey _
As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" _
Alias "RegDeleteValueA" (ByVal HKEY As Long, ByVal _
lpValueName As String) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" _
Alias "RegOpenKeyA" (ByVal HKEY As Long, ByVal lpSubKey _
As String, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, lpData _
As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" _
Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal _
lpValueName As String, ByVal Reserved As Long, ByVal _
dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Public Function GetSettingString(HKEY As Long, _
strPath As String, strValue As String, Optional _
Default As String) As String
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

'Set up default value
If Not IsEmpty(Default) Then
GetSettingString = Default
Else
GetSettingString = ""
End If

lRegResult = RegOpenKey(HKEY, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_SZ Then

strBuffer = String(lDataBufferSize, " ")
lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, _
ByVal strBuffer, lDataBufferSize)

intZeroPos = InStr(strBuffer, Chr$(0))
If intZeroPos > 0 Then
GetSettingString = Left$(strBuffer, intZeroPos - 1)
Else
GetSettingString = strBuffer
End If

End If

Else
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function
Public Function GetSettingLong(ByVal HKEY As Long, _
ByVal strPath As String, ByVal strValue As String, _
Optional Default As Long) As Long

Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long

'Set up default value
If Not IsEmpty(Default) Then
GetSettingLong = Default
Else
GetSettingLong = 0
End If

lRegResult = RegOpenKey(HKEY, strPath, hCurKey)
lDataBufferSize = 4 '4 bytes = 32 bits = long

lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, lBuffer, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_DWORD Then
GetSettingLong = lBuffer
End If

Else
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingLong(ByVal HKEY As Long, ByVal _
strPath As String, ByVal strValue As String, ByVal lData As Long)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(HKEY, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0&, _
REG_DWORD, lData, 4)

If lRegResult <> ERROR_SUCCESS Then
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub
Public Sub SaveSettingString(HKEY As Long, strPath _
As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(HKEY, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, _
ByVal strData, Len(strData))

If lRegResult <> ERROR_SUCCESS Then
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingByte(ByVal HKEY As Long, _
ByVal strPath As String, ByVal strValueName As String, _
Optional Default As Variant) As Variant
Dim lValueType As Long
Dim byBuffer() As Byte
Dim lDataBufferSize As Long
Dim lRegResult As Long
Dim hCurKey As Long

If Not IsEmpty(Default) Then
If VarType(Default) = vbArray + vbByte Then
GetSettingByte = Default
Else
GetSettingByte = 0
End If

Else
GetSettingByte = 0
End If

lRegResult = RegOpenKey(HKEY, strPath, hCurKey)

lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, _
lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_BINARY Then

ReDim byBuffer(lDataBufferSize - 1) As Byte
lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, _
lValueType, byBuffer(0), lDataBufferSize)
GetSettingByte = byBuffer

End If

Else
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Function


Public Function GetAllValues(HKEY As Long, _
strPath As String) As Variant
'Returns: a 2D array.
'(x,0) is value name
'(x,1) is value type (see constants)

Dim lRegResult As Long
Dim hCurKey As Long
Dim lValueNameSize As Long
Dim strValueName As String
Dim lCounter As Long
Dim byDataBuffer(4000) As Byte
Dim lDataBufferSize As Long
Dim lValueType As Long
Dim strNames() As String
Dim lTypes() As Long
Dim intZeroPos As Integer

lRegResult = RegOpenKey(HKEY, strPath, hCurKey)

Do
    'Initialise bufffers
    lValueNameSize = 255
    strValueName = String$(lValueNameSize, " ")
    lDataBufferSize = 4000
    
    lRegResult = RegEnumValue(hCurKey, lCounter, _
    strValueName, lValueNameSize, 0&, lValueType, _
    byDataBuffer(0), lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
    
        'Save the type
        ReDim Preserve strNames(lCounter) As String
        ReDim Preserve lTypes(lCounter) As Long
        lTypes(UBound(lTypes)) = lValueType
        
        'Tidy up string and save it
        intZeroPos = InStr(strValueName, Chr$(0))
        If intZeroPos > 0 Then
            strNames(UBound(strNames)) = _
            Left$(strValueName, intZeroPos - 1)
        Else
            strNames(UBound(strNames)) = strValueName
        End If
    
        lCounter = lCounter + 1
    Else
        Exit Do
    End If
Loop

'Move data into array
Dim Finisheddata() As Variant
On Error GoTo GetOut
ReDim Finisheddata(UBound(strNames), 0 To 1) As Variant

For lCounter = 0 To UBound(strNames)
    Finisheddata(lCounter, 0) = strNames(lCounter)
    Finisheddata(lCounter, 1) = lTypes(lCounter)
Next
GetOut:
Err.Clear
On Error GoTo 0
GetAllValues = Finisheddata

End Function

Public Function GetAllKeys(HKEY As Long, _
strPath As String) As Variant
Dim lRegResult As Long
Dim lCounter As Long
Dim hCurKey As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim strNames() As String
Dim intZeroPos As Integer
lCounter = 0
lRegResult = RegOpenKey(HKEY, strPath, hCurKey)

Do
'initialise buffers (longest possible length=255)
lDataBufferSize = 255
strBuffer = String(lDataBufferSize, " ")
lRegResult = RegEnumKey(hCurKey, _
lCounter, strBuffer, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

'tidy up string and save it
ReDim Preserve strNames(lCounter) As String

intZeroPos = InStr(strBuffer, Chr$(0))
If intZeroPos > 0 Then
strNames(UBound(strNames)) = Left$(strBuffer, intZeroPos - 1)
Else
strNames(UBound(strNames)) = strBuffer
End If

lCounter = lCounter + 1
Else
Exit Do
End If
Loop
GetAllKeys = strNames
End Function


Public Sub SaveSettingByte(ByVal HKEY As Long, ByVal _
strPath As String, ByVal strValueName As String, byData() As Byte)
Dim lRegResult As Long
Dim hCurKey As Long

lRegResult = RegCreateKey(HKEY, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValueName, _
0&, REG_BINARY, byData(0), UBound(byData()) + 1)

lRegResult = RegCloseKey(hCurKey)
End Sub
Function DirExist(ByVal pstrDirName As String) As Boolean
    If Right(pstrDirName, 1) = "\" Then pstrDirName = Left(pstrDirName, Len(pstrDirName) - 1)
    On Error Resume Next
    DirExist = ((GetAttr(pstrDirName) And vbDirectory) = vbDirectory)
End Function
Function IsArrayEmpty(sIN As Variant) As Boolean
On Error GoTo EmptyArray
Dim x As Integer
    IsArrayEmpty = False
    x = UBound(sIN)
    On Error GoTo 0
    Exit Function
EmptyArray:
    On Error GoTo 0
    IsArrayEmpty = True
End Function
