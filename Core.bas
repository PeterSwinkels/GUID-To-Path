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
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
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
Private Const NO_KEY As Long = 0&                 'Defines a null registry key.


'This procedure manages/returns the registry key access mode used.
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

'This procedure checks the key at the specified path for the specified GUID of the specified type and returns the result.
Private Function CheckKeyPath(HiveH As Long, KeyPath As String, GUID As String, GUIDType As String, HiveKeyName As String) As String
On Error GoTo ErrorTrap
Dim GUIDParentKeyH As Long
Dim Result As String
   
   GUIDParentKeyH = OpenKeyPath(HiveH, KeyPath)
   If Not GUIDParentKeyH = NO_KEY Then
      Result = Result & GetGUIDProperties(GUIDParentKeyH, GUID, GUIDType, HiveKeyName)
      RegCloseKey GUIDParentKeyH
   End If
   
EndRoutine:
   CheckKeyPath = Result
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
   Else
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
Dim BitModeFlag As Long
Dim BitModeFlags() As Variant
Dim GUIDType As Long
Dim GUIDTypes() As Variant
Dim HiveKey As Long
Dim HiveKeys() As Variant
Dim KeyH As Long
Dim Result As String
Dim ReturnValue As Long

   BitModeFlags = Array(False, True)
   GUIDTypes = Array("AppID", "CLSID", "Component Categories", "Interface", "TypeLib")
   HiveKeys = Array("HKCR", "HKCU", "HKLM")
   Result = vbNullString
   If Not GUID = vbNullString Then
      For BitModeFlag = LBound(BitModeFlags()) To UBound(BitModeFlags())
         For HiveKey = LBound(HiveKeys()) To UBound(HiveKeys())
            AccessMode NewIs64Bit:=CBool(BitModeFlags(BitModeFlag))
            
            For GUIDType = LBound(GUIDTypes()) To UBound(GUIDTypes())
               Select Case CStr(HiveKeys(HiveKey))
                  Case "HKCR"
                     Result = Result & CheckKeyPath(HKEY_CLASSES_ROOT, CStr(GUIDTypes(GUIDType)), GUID, CStr(GUIDTypes(GUIDType)), CStr(HiveKeys(HiveKey)))
                  Case "HKCU"
                     Result = Result & CheckKeyPath(HKEY_CURRENT_USER, "SOFTWARE\Classes\" & CStr(GUIDTypes(GUIDType)), GUID, CStr(GUIDTypes(GUIDType)), CStr(HiveKeys(HiveKey)))
               End Select
            Next GUIDType
            
            Select Case CStr(HiveKeys(HiveKey))
               Case "HKLM"
                  Result = Result & CheckKeyPath(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions", GUID, vbNullString, "HKLM")
                  Result = Result & CheckKeyPath(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes", GUID, vbNullString, "HKLM")
                  Result = Result & CheckKeyPath(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Class", GUID, vbNullString, "HKLM")
            End Select
            
         Next HiveKey
      Next BitModeFlag
      
      If Result = vbNullString Then Result = GUID & " - not found." & vbCrLf
      
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
Public Function FormatGUID(GUID As String) As String
On Error GoTo ErrorTrap
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

'This procedure returns the first GUID in the specified text if present.
Public Function GetGUIDFromText(Text As String) As String
On Error GoTo ErrorTrap
Dim GUID As String
Dim EndPosition As String
Dim StartPosition As String

   GUID = vbNullString
   StartPosition = InStr(Text, "{")
   If StartPosition > 0 Then
      EndPosition = InStr(StartPosition + 1, Text, "}")
      If EndPosition > 0 Then
         GUID = Mid$(Text, StartPosition, (EndPosition - StartPosition) + 1)
      End If
   End If

EndRoutine:
   GetGUIDFromText = GUID
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the properties for the specified GUID of the specified type.
Private Function GetGUIDProperties(KeyH As Long, GUID As String, GUIDType As String, HiveKeyName As String) As String
On Error GoTo ErrorTrap
Dim GUIDKeyH As Long
Dim Paths As String
Dim Result As String
Dim ReturnValue As Long

   ReturnValue = RegOpenKeyExA(KeyH, GUID, CLng(0), AccessMode(), GUIDKeyH)
   If ReturnValue = ERROR_SUCCESS Then
      Result = Result & GUID & " (" & HiveKeyName & ") (" & GUIDType & ") "
      
      If Is64BitAccess(AccessMode()) Then
         Result = Result & "(64 bit)" & vbCrLf
      Else
         Result = Result & "(32 bit)" & vbCrLf
      End If
      
      Paths = GetPathsFromGUID(GUIDKeyH, GUID)
      If Paths = vbNullString Then
         Result = Result & "No handler/server paths." & vbCrLf
      Else
         Result = Result & Paths & vbCrLf
      End If
      
      Result = Result & GetRegistryValueAsText(KeyH, GUID, "CanonicalName", "Canonical name")
      Result = Result & GetRegistryValueAsText(KeyH, GUID, vbNullString, "Default")
      Result = Result & GetRegistryValueAsText(KeyH, GUID, "Class", "Class")
      Result = Result & GetRegistryValueAsText(KeyH, GUID, "Name", "Name")
      Result = Result & vbCrLf
      
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

   Index = 0
   Do Until ReturnValue = ERROR_NO_MORE_ITEMS Or (Not ReturnValue = ERROR_SUCCESS)
      KeyName = String$(MAX_SHORT_STRING, vbNullChar)
      Length = Len(KeyName)
      ReturnValue = RegEnumKeyExA(ParentKeyH, Index, KeyName, Length, CLng(0), vbNullString, CLng(0), CLng(0))
      If ReturnValue = ERROR_SUCCESS Then
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
               If Not SafeArrayGetDim(SubKeys()) = 0 Then
                  For Each SubKey In SubKeys
                     If IsWholeNumber(CStr(SubKey)) Then
                        ReturnValue = RegOpenKeyExA(KeyH, CStr(SubKey), CLng(0), AccessMode(), SubKeyH)
                        If ReturnValue = ERROR_SUCCESS Then
                           For Each KeyName In Array("Win32", "Win64")
                              Value = GetRegistryValue(SubKeyH, CStr(KeyName), vbNullString)
                              If Not Value = vbNullString Then
                                 Result = Result & KeyName & " = """ & Value & """" & vbCrLf
                              End If
                           Next KeyName
                           RegCloseKey SubKeyH
                        End If
                     End If
                  Next SubKey
               End If
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

      If ReturnValue = ERROR_SUCCESS And Length > 0 Then
         ValueData = Left$(ValueData, Length - 1)
      Else
         ValueData = vbNullString
      End If
   
      RegCloseKey KeyH
   End If
   
EndRoutine:
   GetRegistryValue = ValueData
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure attempts to retrieve the specified registry value and returns it formatted as text if found.
Private Function GetRegistryValueAsText(KeyH As Long, GUID As String, ValueName As String, Description As String) As String
On Error GoTo ErrorTrap
Dim Result As String
Dim Value As String

   Value = GetRegistryValue(KeyH, GUID, ValueName)
   If Not Value = vbNullString Then
      Result = Description & " = """ & Value & """" & vbCrLf
   End If
   
EndRoutine:
   GetRegistryValueAsText = Result
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
   If MsgBox(Description, vbOKCancel Or vbDefaultButton1 Or vbExclamation) = vbCancel Then
      Resume EndProgram
   End If
   Exit Sub

EndProgram:
   End
   
ErrorTrap:
   Resume EndProgram
End Sub

'This procedure checks whether the specified mode indicates 64 bit access and returns the result.
Private Function Is64BitAccess(Mode As Long) As Boolean
On Error GoTo ErrorTrap
Dim Is64Bit As Boolean

   Is64Bit = ((Mode And KEY_WOW64_64KEY) = KEY_WOW64_64KEY)

EndRoutine:
   Is64BitAccess = Is64Bit
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

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


'This procedure returns the handle for the last key in the specified path.
Private Function OpenKeyPath(HiveH As Long, KeyPath As String) As Long
On Error GoTo ErrorTrap
Dim KeyH As Long
Dim KeyNames As Variant
Dim Index As Long
Dim ResultKeyH As Long
Dim ReturnValue As Long
Dim SubKeyH As Long

   KeyH = HiveH
   KeyNames = Split(KeyPath, "\")
   Index = LBound(KeyNames, 1)
   ResultKeyH = NO_KEY
   Do
      ReturnValue = RegOpenKeyExA(KeyH, CStr(KeyNames(Index)), CLng(0), AccessMode(), SubKeyH)
      If ReturnValue = ERROR_SUCCESS Then
         If Index = UBound(KeyNames, 1) Then
            ResultKeyH = SubKeyH
         Else
            RegCloseKey KeyH
            KeyH = SubKeyH
            Index = Index + 1
         End If
      End If
   Loop While ResultKeyH = NO_KEY And ReturnValue = ERROR_SUCCESS

EndRoutine:
   OpenKeyPath = ResultKeyH
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


