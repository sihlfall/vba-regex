Attribute VB_Name = "MkConsoleApp"
' ------------------------------------------------------
' MakeConsole.BAS -- Copyright (c) Slightly Tilted Software
' By: L.J. Johnson       Date: 11-30-1997
' Comments:    Contains MAIN(), plus the function
'              which take a standard VB 5.0 EXE
'              and change it to a 32-bit console app
' ------------------------------------------------------
Option Explicit
Option Base 1
DefLng A-Z

Private Const GENERIC_READ                As Long = &H80000000
Private Const OPEN_EXISTING               As Long = 3&
Private Const FILE_ATTRIBUTE_NORMAL       As Long = &H80&

Private Const SCS_32BIT_BINARY = 0&
Private Const SCS_DOS_BINARY = 1&
Private Const SCS_WOW_BINARY = 2&
Private Const SCS_PIF_BINARY = 3&
Private Const SCS_POSIX_BINARY = 4&
Private Const SCS_OS216_BINARY = 5&

Private Const constMsgTitle = "Make Console App"

' ---------------------------------------------
' Windows API calls
' ---------------------------------------------
Public Declare Sub CopyMem _
   Lib "kernel32" Alias "RtlMoveMemory" _
   (dst As Any, src As Any, ByVal Size As Long)
Private Declare Function CloseHandle _
   Lib "kernel32" _
   (ByVal hObject As Long) As Long
Private Declare Function CreateFile _
   Lib "kernel32" Alias "CreateFileA" _
   (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Public Sub Main()
   Dim strCmd              As String
   Dim strMsg              As String
   Dim strRtn              As String
   
   strCmd = Command$
   
   If Trim$(strCmd) = "" Then
      strMsg = "You must enter the name of a VB 5.0 standard executable file."
      MsgBox strMsg, vbExclamation, constMsgTitle
   Else
      If InStr(1, strCmd, ".", vbTextCompare) = 0 Then
         strCmd = strCmd & ".EXE"
      End If
      
      If Exists(strCmd) = True Then
         strRtn = SetConsoleApp(strCmd)
         MsgBox strRtn, vbInformation, constMsgTitle
      Else
         strMsg = "The file, " & Trim$(strCmd) & ", does not exist."
         MsgBox strMsg, vbCritical, constMsgTitle
      End If
   End If
   
End Sub

Private Function SetConsoleApp(xstrFileName As String) As String
   Dim lngFileNum          As Long
   Dim ststrMZ_Header      As String * 512
   Dim strMagic            As String * 2
   Dim strMagicPE          As String * 2
   Dim lngNewPE_Offset     As Long
   Dim lngData             As Long
   Dim strTmp              As String
   Const PE_FLAG_OFFSET    As Long = 93&
   Const DOS_FILE_OFFSET   As Long = 25&
   
   ' ---------------------------------------------
   ' See if file actually exists
   ' ---------------------------------------------
   strTmp = Trim$(Dir$(xstrFileName))
   If Len(strTmp) = 0 Then
      SetConsoleApp = "Failed -- The file, " & xstrFileName & ", does not exist!"
      GoTo ExitCheck
   End If
       
   ' ---------------------------------------------
   ' Get a free file handle
   ' ---------------------------------------------
   On Error Resume Next
   lngFileNum = FreeFile
   Open xstrFileName For Binary Access Read Write Shared As lngFileNum
   
   ' ---------------------------------------------
   ' Get the first 512 characters from from file
   ' ---------------------------------------------
   Seek #lngFileNum, 1
   Get lngFileNum, , ststrMZ_Header
   
   ' ---------------------------------------------
   ' Look for the "magic header" values "MZ"
   ' If it doesn't exist, then it's not an EXE file
   ' ---------------------------------------------
   If Mid$(ststrMZ_Header, 1, 2) <> "MZ" Then
      SetConsoleApp = "Failed -- File is not an executable file."
      GoTo ExitCheck
   End If
   
   ' ---------------------------------------------
   ' Check to see if it's a MS-DOS executable
   ' ---------------------------------------------
   CopyMem lngData, ByVal Mid$(ststrMZ_Header, DOS_FILE_OFFSET, 2), 2
   If lngData < 64 Then
      SetConsoleApp = "Failed -- File is 16-bit MSDOS EXE file."
      GoTo ExitCheck
   End If
   
   ' ---------------------------------------------
   ' Get the offset for the new .EXE header
   ' ---------------------------------------------
   CopyMem lngNewPE_Offset, ByVal Mid$(ststrMZ_Header, 61, 4), 4
   
   ' ---------------------------------------------
   ' Get the "magic" header (NE, LE, PE)
   ' ---------------------------------------------
   strMagic = Mid$(ststrMZ_Header, lngNewPE_Offset + 1, 2)
   strMagicPE = Mid$(ststrMZ_Header, lngNewPE_Offset + 3, 2)
   
   Select Case strMagic
      
      ' ---------------------------------------------
      ' Check for NT format
      ' ---------------------------------------------
      Case "PE"
         If strMagicPE <> vbNullChar & vbNullChar Then
            SetConsoleApp = "Failed -- File is unknown 32-bit NT executable file."
            GoTo ExitCheck
         End If
         
         ' ---------------------------------------------
         ' Get the subsystem flags to identify NT
         '     character-mode
         ' ---------------------------------------------
         lngData = Asc(Mid$(ststrMZ_Header, lngNewPE_Offset + PE_FLAG_OFFSET, 1))
         If lngData <> 3 Then
            On Error Resume Next
            Err.Number = 0
            Seek #lngFileNum, lngNewPE_Offset + PE_FLAG_OFFSET
            Put lngFileNum, , 3
            If Err.Number = 0 Then
               SetConsoleApp = "Success -- Converted file to console app."
            Else
               SetConsoleApp = "Failed -- Error converting to console app: " & Err.Description
            End If
         Else
            SetConsoleApp = "Failed -- Already a console app"
         End If
         
      Case Else
         SetConsoleApp = "Failed -- Not correct file type."
         
   End Select

ExitCheck:
   ' ---------------------------------------------
   ' Close the file
   ' ---------------------------------------------
   Close lngFileNum
   
   On Error GoTo 0
   
End Function

Public Function Exists(ByVal xstrFullName As String) As Boolean
On Error Resume Next       ' Don't accept errors here
   Const constProcName     As String = "Exists"
   Dim lngFileHwnd         As Long
   Dim lngRtn              As Long

   ' ------------------------------------------
   ' Open the file only if it already exists
   ' ------------------------------------------
   lngFileHwnd = CreateFile(xstrFullName, _
                            GENERIC_READ, 0&, _
                            0&, OPEN_EXISTING, _
                            FILE_ATTRIBUTE_NORMAL, 0&)
   
   ' ------------------------------------------
   ' If get these specific errors, then
   '     file doesn't exist
   ' ------------------------------------------
   If lngFileHwnd = 0 Or lngFileHwnd = -1 Then
      Exists = False
   Else
      ' Success -- Must close the handle
      lngRtn = CloseHandle(lngFileHwnd)
      Exists = True
   End If

On Error GoTo 0
End Function


