Attribute VB_Name = "SysInfo"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Returns True if the version of Windows that the user is running is greater
' or equal than Vista (including 7, 8, ...)
Public Function IsWindowsVistaOrGreater() As Boolean
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                IsWindowsVistaOrGreater = False   ' Windows 3.1
            Case VER_PLATFORM_WIN32_NT
                 Select Case osv.dwVerMajor
                    Case 3
                        IsWindowsVistaOrGreater = False  ' NT 3.5
                    Case 4
                        IsWindowsVistaOrGreater = False  ' NT 4.0
                    Case 5
                        IsWindowsVistaOrGreater = False  ' 2000, XP
                    Case 6
                        IsWindowsVistaOrGreater = True   ' Vista, 7, ...
                End Select
            Case VER_PLATFORM_WIN32_WINDOWS:
                IsWindowsVistaOrGreater = False   '95, Me or 98
        End Select
    Else
        IsWindowsVistaOrGreater = False
    End If
End Function
