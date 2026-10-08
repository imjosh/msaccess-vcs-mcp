Attribute VB_Name = "modSideRegistry"
Option Explicit

' This module belongs only to the disposable user database. The loaded library
' is never modified. HKCU remapping affects this Access process only.
Private Declare PtrSafe Function RegCreateKeyExW Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr, ByVal subKey As LongPtr, ByVal reserved As Long, _
    ByVal className As LongPtr, ByVal options As Long, ByVal access As Long, _
    ByVal security As LongPtr, ByRef result As LongPtr, ByRef disposition As Long) As Long
Private Declare PtrSafe Function RegOverridePredefKey Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr, ByVal newKey As LongPtr) As Long
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long

Public Function SideRegistryBegin(ByVal path As String) As Long
    Dim key As LongPtr
    Dim disposition As Long
    Dim status As Long
    ' SaveSetting creates ordinary subkeys, which Windows disallows below a
    ' volatile key. Use an ordinary, unique temporary tree and delete on exit.
    status = RegCreateKeyExW(&H80000001, StrPtr(path), 0, 0, 0, &HF003F, 0, key, disposition)
    If status <> 0 Then SideRegistryBegin = status: Exit Function
    If disposition <> 1 Then
        RegCloseKey key
        SideRegistryBegin = 183
        Exit Function
    End If
    SideRegistryBegin = RegOverridePredefKey(&H80000001, key)
    RegCloseKey key
End Function

Public Function SideRegistryEnd(ByVal ignored As String) As Long
    SideRegistryEnd = RegOverridePredefKey(&H80000001, 0)
End Function
