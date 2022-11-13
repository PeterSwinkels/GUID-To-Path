Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants and functions used by this program.
Private Const ERROR_FILE_NOT_FOUND As Long = 2&
Private Const ERROR_NO_MORE_ITEMS As Long = 259&
Private Const ERROR_SUCCESS As Long = 0&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const KEY_READ As Long = &H20019
Private Const KEY_WOW64_64KEY As Long = &H100&
Private Const MAX_REG_VALUE_DATA As Long = &HFFFFF
Private Const REG_SZ As Long = &H1&

Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Long) As Long
Private Declare Function RegOpenKeyExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long

'The constants used by this program.
Private Const MAX_LONG_STRING As Long = &HFFFF&   'Defines the maximum length in bytes allowed for a long string.
Private Const MAX_SHORT_STRING As Long = &HFF&    'Defines the maximum length in bytes allowed for a short string.


'This procedure manages the access mode used.
Private Function AccessMode(Optional NewIs64Bit As Variant) As Long
On Error GoTo ErrorTrap
Dim Mode As Long
Static CurrentIs64Bit As Boolean

   If Not IsMissing(NewIs64Bit) Then CurrentIs64Bit = CBool(NewIs64Bit)

   Mode = KEY_READ
   If CurrentIs64Bit Then Mode = Mode Or KEY_WOW64_64KEY

EndRoutine:
   AccessMode = Mode
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure checks the HKEY_CLASSES_ROOT key for the specified GUID of the specified type and returns the result.
Private Function CheckHKEYCLASSESROOT(GUID As String, GUIDType As String, ByRef Found As Boolean) As String
On Error GoTo ErrorTrap
Dim KeyH As Long
Dim Result As String
Dim ReturnValue As Long

   ReturnValue = RegOpenKeyExA(HKEY_CLASSES_ROOT, GUIDType, CLng(0), AccessMode(), KeyH)
   
   If ReturnValue = ERROR_SUCCESS Then
      Result = Result & GetGUIDProperties(KeyH, GUID, GUIDType, Found)
      RegCloseKey KeyH
   End If
   
EndRoutine:
   CheckHKEYCLASSESROOT = Result
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the description for the specified error code.
Private Function ErrorDescription(ErrorCode As Long) As String
On Error GoTo ErrorTrap
Dim Description As String
Dim Length As Long

   Description = String$(MAX_LONG_STRING, vbNullChar)
   Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
   If Length = 0 Then
      Description = "No description."
   ElseIf Length > 0 Then
      Description = Left$(Description, Length - 1)
   End If
   
EndRoutine:
   ErrorDescription = Description
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure searches for the specified GUID and gives the command to retrieve any paths found.
Public Function FindGUID(GUID As String) As String
On Error GoTo ErrorTrap
Dim Found As Boolean
Dim GUIDType As Long
Dim GUIDTypes() As Variant
Dim KeyH As Long
Dim Result As String
Dim ReturnValue As Long

   AccessMode NewIs64Bit:=False
   Result = vbNullString
   If Not GUID = vbNullString Then
      Do While DoEvents() > 0
         Found = False
         GUIDTypes = Array("AppID", "CLSID", "Interface", "TypeLib")
         For GUIDType = LBound(GUIDTypes()) To UBound(GUIDTypes())
            Result = Result & CheckHKEYCLASSESROOT(GUID, CStr(GUIDTypes(GUIDType)), Found)
         Next GUIDType
         
         If Found Then
            Exit Do
         Else
            If AccessMode() = (KEY_READ Or KEY_WOW64_64KEY) Then
               Result = Result & GUID & " (?)" & vbCrLf
               Result = Result & "GUID not found." & vbCrLf
               Exit Do
            Else
               Result = Result & "[Attempting 64 bit mode.]" & vbCrLf
               AccessMode NewIs64Bit:=True
            End If
         End If
      Loop
      
      Result = Result & vbCrLf
   End If
   
EndRoutine:
   FindGUID = Result
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure formats the specified GUID and returns the result.
Public Function FormatGUID(GUID As String)
Dim FormattedGUID As String

   FormattedGUID = UCase$(Trim$(GUID))
   
   If Not FormattedGUID = vbNullString Then
      If Not Left$(FormattedGUID, 1) = "{" Then FormattedGUID = "{" & FormattedGUID
      If Not Right$(FormattedGUID, 1) = "}" Then FormattedGUID = FormattedGUID & "}"
   End If
   
EndRoutine:
   FormatGUID = FormattedGUID
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function
'This procedure returns the properties for the specified GUID of the specified type.
Private Function GetGUIDProperties(KeyH As Long, GUID As String, GUIDType As String, ByRef Found As Boolean) As String
On Error GoTo ErrorTrap
Dim GUIDKeyH As Long
Dim Paths As String
Dim Result As String
Dim ReturnValue As Long

   ReturnValue = RegOpenKeyExA(KeyH, GUID, CLng(0), AccessMode(), GUIDKeyH)

   If ReturnValue = ERROR_SUCCESS Then
      Result = Result & GUID & " (" & GUIDType & ")" & vbCrLf
      Result = Result & "Default = """ & GetRegistryValue(KeyH, GUID, vbNullString) & """" & vbCrLf

      Found = True
      Paths = GetPathsFromGUID(GUIDKeyH, GUID)
      If Paths = vbNullString Then Result = Result & "No paths." & vbCrLf Else Result = Result & Paths
      
      RegCloseKey GUIDKeyH
   ElseIf Not ReturnValue = ERROR_FILE_NOT_FOUND Then
      Result = Result & "Error code: " & CStr(ReturnValue) & " - """ & ErrorDescription(ReturnValue) & """" & vbCrLf
   End If

EndRoutine:
   GetGUIDProperties = Result
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function



'This procedure returns the registry keys contained by the specified key.
Private Function GetKeys(ParentKeyH As Long) As String()
On Error GoTo ErrorTrap
Dim Index As Long
Dim KeyH As Long
Dim KeyName As String
Dim Keys() As String
Dim Length As Long
Dim ReturnValue As Long

   Do
      KeyName = String$(MAX_SHORT_STRING, vbNullChar)
      Length = Len(KeyName)
      ReturnValue = RegEnumKeyExA(ParentKeyH, Index, KeyName, Length, CLng(0), vbNullString, CLng(0), CLng(0))
      If ReturnValue = ERROR_NO_MORE_ITEMS Or Not ReturnValue = ERROR_SUCCESS Then
         Exit Do
      Else
         If SafeArrayGetDim(Keys()) = 0 Then
            ReDim Keys(0 To 0) As String
         Else
            ReDim Preserve Keys(LBound(Keys()) To UBound(Keys()) + 1) As String
         End If
   
         Keys(UBound(Keys())) = Left$(KeyName, Length)
         Index = Index + 1
      End If
   Loop
   
EndRoutine:
   GetKeys = Keys()
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure retrieves the paths referred to by the specified GUID.
Private Function GetPathsFromGUID(GUIDKeyH As Long, GUID As String) As String
On Error GoTo ErrorTrap
Dim Key As Variant
Dim KeyH As Long
Dim KeyName As Variant
Dim Keys() As String
Dim Result As String
Dim ReturnValue As Long
Dim SubKey As Variant
Dim SubKeyH As Long
Dim SubKeys() As String
Dim Value As String

   For Each KeyName In Array("InprocHandler", "InprocHandler32", "InprocServer", "InprocServer32", "LocalServer", "LocalServer32")
      Value = GetRegistryValue(GUIDKeyH, CStr(KeyName), vbNullString)
      If Not Value = vbNullString Then Result = Result & KeyName & " = """ & Value & """" & vbCrLf
   Next KeyName
   
   For Each KeyName In Array("ProxyStubClsid", "ProxyStubClsid32")
      Value = UCase$(Trim$(GetRegistryValue(GUIDKeyH, CStr(KeyName), vbNullString)))
      If Not Value = vbNullString Then
         If Not Value = GUID Then Result = Result & KeyName & " = " & FindGUID(Value)
      End If
   Next KeyName
   
   Keys = GetKeys(GUIDKeyH)
   If Not SafeArrayGetDim(Keys()) = 0 Then
      For Each Key In Keys
         If IsVersion(CStr(Key)) Then
            ReturnValue = RegOpenKeyExA(GUIDKeyH, CStr(Key), CLng(0), AccessMode(), KeyH)
            If ReturnValue = ERROR_SUCCESS Then
               SubKeys = GetKeys(KeyH)
               For Each SubKey In SubKeys
                  If IsWholeNumber(CStr(SubKey)) Then
                     ReturnValue = RegOpenKeyExA(KeyH, CStr(SubKey), CLng(0), AccessMode(), SubKeyH)
                     If ReturnValue = ERROR_SUCCESS Then
                        For Each KeyName In Array("Win32", "Win64")
                           Value = GetRegistryValue(SubKeyH, CStr(KeyName), vbNullString)
                           If Not Value = vbNullString Then Result = Result & KeyName & " = """ & Value & """" & vbCrLf
                        Next KeyName
                     End If
                  End If
               Next SubKey
               RegCloseKey KeyH
            End If
         End If
      Next Key
   End If
   
EndRoutine:
   GetPathsFromGUID = Result
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the specified registry value's data.
Private Function GetRegistryValue(ParentKeyH As Long, KeyName As String, ValueName As String) As String
On Error GoTo ErrorTrap
Dim KeyH As Long
Dim Length As Long
Dim ReturnValue As Long
Dim ValueData As String

   ReturnValue = RegOpenKeyExA(ParentKeyH, KeyName, CLng(0), AccessMode(), KeyH)
   If ReturnValue = ERROR_SUCCESS Then
      ValueData = String$(MAX_REG_VALUE_DATA, vbNullChar)
      Length = Len(ValueData)
      ReturnValue = RegQueryValueExA(KeyH, ValueName, CLng(0), REG_SZ, ValueData, Length)

      If ReturnValue = ERROR_SUCCESS Then
         If Length > 0 Then ValueData = Left$(ValueData, Length - 1)
         RegCloseKey KeyH
      Else
         ValueData = vbNullString
      End If
   End If
   
EndRoutine:
   GetRegistryValue = ValueData
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function




'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number
   Err.Clear
   
   On Error GoTo ErrorTrap
   Description = Description & vbCr & "Error code: " & CStr(ErrorCode)
   MsgBox Description, vbExclamation
Exit Sub

EndProgram:
   End
   
ErrorTrap:
   Resume EndProgram
End Sub


'This procedure checks whether the specified value is a version number.
Private Function IsVersion(Value As String) As Boolean
On Error GoTo ErrorTrap
Dim Major As String
Dim Minor As String
Dim Position As Long
Dim Result As Boolean

   Result = False
   Position = InStr(Value, ".")
   If Position > 0 Then
      Major = Left$(Value, Position - 1)
      Minor = Mid$(Value, Position + 1)
      Result = (IsWholeNumber(Major) And IsWholeNumber(Minor))
   End If
EndRoutine:
   IsVersion = Result
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure checks whether specified value is a whole number.
Private Function IsWholeNumber(Value As String) As Boolean
On Error GoTo ErrorTrap
   IsWholeNumber = (Value = CStr(CLng(Val(Value))))
EndRoutine:
   Exit Function
   
ErrorTrap:
   HandleError
   IsWholeNumber = False
   Resume EndRoutine
End Function


'This procedure is started when this program is executed.
Public Sub Main()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path

   InterfaceWindow.Show
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


