Attribute VB_Name = "mdStartup"
'=========================================================================
'
' VbUdpBroadcast (c) 2022 by wqweto@gmail.com
'
' UDP Broadcast Forwarder (based on https://github.com/dechamps/WinIPBroadcast)
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdStartup"

'=========================================================================
' API
'=========================================================================

'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL             As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

'=========================================================================
' Constants and variables
'=========================================================================

Public Const STR_SERVICE_NAME           As String = "VbUdpBroadcast"
'--- messages
Private Const MSG_SVC_INSTALL           As String = "Installing NT service %1"
Private Const MSG_SVC_UNINSTALL         As String = "Uninstalling NT service %1"
Private Const MSG_FAILURE               As String = "Error"
Private Const MSG_SUCCESS               As String = "Success"
Private Const MSG_PREFIX_ERROR          As String = "Error"
Private Const MSG_PREFIX_WARNING        As String = "Warning"
Private Const MSG_PREFIX_DEBUG          As String = "Debug"
'--- formats
Private Const FORMAT_TIME_ONLY           As String = "hh:nn:ss"
Private Const FORMAT_BASE_3              As String = "0.000"
'--- log level
Public Const vbLogEventTypeDebug        As Long = vbLogEventTypeInformation + 1
Public Const vbLogEventTypeDataDump     As Long = vbLogEventTypeInformation + 2

Private m_oOpt                      As Object
Private m_bIsService                As Boolean
Private m_oMain                     As frmMain
Private m_bVerbose                  As Boolean

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get STR_VERSION() As String
    STR_VERSION = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get Verbose() As Boolean
    Verbose = m_bVerbose
End Property

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    Const FUNC_NAME     As String = "Main"
    Dim lExitCode       As Long
    
    On Error GoTo EH
    lExitCode = Process(SplitArgs(Command$))
    If Not InIde And lExitCode <> -1 Then
        Call ExitProcess(lExitCode)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Public Function Process(vArgs As Variant) As Long
    Const FUNC_NAME     As String = "Process"
    Dim vKey            As Variant
    Dim lIdx            As Long
    Dim sError          As String
    
    On Error GoTo EH
    Set m_oOpt = GetOpt(vArgs)
    '--- normalize options: convert -o and -option to proper long form (--option)
    For Each vKey In Split("nologo install:i uninstall:u console help:h:? verbose:v")
        vKey = Split(vKey, ":")
        For lIdx = 0 To UBound(vKey)
            If IsEmpty(m_oOpt.Item("--" & At(vKey, 0))) And Not IsEmpty(m_oOpt.Item("-" & At(vKey, lIdx))) Then
                m_oOpt.Item("--" & At(vKey, 0)) = m_oOpt.Item("-" & At(vKey, lIdx))
            End If
        Next
    Next
    m_bVerbose = C_Bool(m_oOpt.Item("--verbose"))
    If Not C_Bool(m_oOpt.Item("--nologo")) Then
        ConsolePrint App.ProductName & " v" & STR_VERSION & vbCrLf & Replace(App.LegalCopyright, "©", "(c)") & vbCrLf & vbCrLf
    End If
    If C_Bool(m_oOpt.Item("--help")) Then
        ConsolePrint "Usage: " & App.EXEName & ".exe [options...]" & vbCrLf & vbCrLf & _
                    "Options:" & vbCrLf & _
                    "  -i, --install       install NT service" & vbCrLf & _
                    "  -u, --uninstall     remove NT service" & vbCrLf & _
                    "  --console           output to console" & vbCrLf
        GoTo QH
    End If
    If NtServiceInit(STR_SERVICE_NAME) Then
        m_bIsService = True
        Set m_oMain = New frmMain
        '--- cannot handle these as NT service
        m_oOpt.Item("--install") = Empty
        m_oOpt.Item("--uninstall") = Empty
        m_oOpt.Item("--console") = True
    End If
    If C_Bool(m_oOpt.Item("--install")) Then
        ConsolePrint Printf(MSG_SVC_INSTALL, STR_SERVICE_NAME) & vbCrLf
        If Not NtServiceInstall(STR_SERVICE_NAME, App.ProductName & " (" & STR_VERSION & ")", GetProcessName(), Error:=sError) Then
            ConsoleError MSG_FAILURE & ": "
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sError & vbCrLf
        Else
            ConsolePrint MSG_SUCCESS & vbCrLf
        End If
        GoTo QH
    ElseIf C_Bool(m_oOpt.Item("--uninstall")) Then
        ConsolePrint Printf(MSG_SVC_UNINSTALL, STR_SERVICE_NAME) & vbCrLf
        If Not NtServiceUninstall(STR_SERVICE_NAME, Error:=sError) Then
            ConsoleError MSG_FAILURE & ": "
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sError & vbCrLf
        Else
            ConsolePrint MSG_SUCCESS & vbCrLf
        End If
        GoTo QH
    End If
    If C_Bool(m_oOpt.Item("--console")) Then
        Process = -1
        Set m_oMain = New frmMain
    End If
    If Not m_oMain Is Nothing Then
        If Not m_oMain.Init() Then
            GoTo QH
        End If
    End If
    If m_bIsService Then
        Do While Not NtServiceQueryStop()
            '--- do nothing
        Loop
        m_oMain.Terminate
        NtServiceTerminate
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function SplitArgs(sText As String) As Variant
    Dim vRetVal         As Variant
    Dim lPtr            As Long
    Dim lArgc           As Long
    Dim lIdx            As Long
    Dim lArgPtr         As Long

    If LenB(sText) <> 0 Then
        lPtr = CommandLineToArgvW(StrPtr(sText), lArgc)
    End If
    If lArgc > 0 Then
        ReDim vRetVal(0 To lArgc - 1) As String
        For lIdx = 0 To UBound(vRetVal)
            Call CopyMemory(lArgPtr, ByVal lPtr + 4 * lIdx, 4)
            vRetVal(lIdx) = SysAllocString(lArgPtr)
        Next
    Else
        vRetVal = Split(vbNullString)
    End If
    Call LocalFree(lPtr)
    SplitArgs = vRetVal
End Function

Private Function SysAllocString(ByVal lPtr As Long) As String
    Dim lTemp           As Long

    lTemp = ApiSysAllocString(lPtr)
    Call CopyMemory(ByVal VarPtr(SysAllocString), lTemp, 4)
End Function

Public Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Public Function GetOpt(vArgs As Variant, Optional OptionsWithArg As String) As Object
    Dim oRetVal         As Object
    Dim lIdx            As Long
    Dim bNoMoreOpt      As Boolean
    Dim vOptArg         As Variant
    Dim vElem           As Variant

    vOptArg = Split(OptionsWithArg, ":")
    Set oRetVal = CreateObject("Scripting.Dictionary")
    With oRetVal
        .CompareMode = vbTextCompare
        For lIdx = 0 To UBound(vArgs)
            Select Case Left$(At(vArgs, lIdx), 1 + bNoMoreOpt)
            Case "-", "/"
                For Each vElem In vOptArg
                    If Mid$(At(vArgs, lIdx), 2, Len(vElem)) = vElem Then
                        If Mid(At(vArgs, lIdx), Len(vElem) + 2, 1) = ":" Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 3)
                        ElseIf Len(At(vArgs, lIdx)) > Len(vElem) + 1 Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 2)
                        ElseIf LenB(At(vArgs, lIdx + 1)) <> 0 Then
                            .Item("-" & vElem) = At(vArgs, lIdx + 1)
                            lIdx = lIdx + 1
                        Else
                            .Item("error") = "Option -" & vElem & " requires an argument"
                        End If
                        GoTo Continue
                    End If
                Next
                .Item("-" & Mid$(At(vArgs, lIdx), 2)) = True
            Case Else
                .Item("numarg") = .Item("numarg") + 1
                .Item("arg" & .Item("numarg")) = At(vArgs, lIdx)
            End Select
Continue:
        Next
    End With
    Set GetOpt = oRetVal
End Function

Public Function C_Str(Value As Variant) As String
    Dim vDest           As Variant

    If VarType(Value) = vbString Then
        C_Str = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbString) = 0 Then
        C_Str = vDest
    End If
End Function

Public Function C_Bool(Value As Variant) As Boolean
    Dim vDest           As Variant

    If VarType(Value) = vbBoolean Then
        C_Bool = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbBoolean) = 0 Then
        C_Bool = vDest
    End If
End Function

Public Property Get At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error GoTo QH
    At = sDefault
    If IsArray(vData) Then
        If lIdx < LBound(vData) Then
            '--- lIdx = -1 for last element
            lIdx = UBound(vData) + 1 + lIdx
        End If
        If LBound(vData) <= lIdx And lIdx <= UBound(vData) Then
            At = C_Str(vData(lIdx))
        End If
    End If
QH:
End Property

Public Function Printf(ByVal sText As String, ParamArray A() As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    
    For lIdx = UBound(A) To LBound(A) Step -1
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE)))
    Next
    Printf = Replace(sText, ChrW$(LNG_PRIVATE), "%")
End Function

Public Function GetProcessName() As String
    GetProcessName = String$(1000, 0)
    Call GetModuleFileName(0, GetProcessName, Len(GetProcessName) - 1)
    GetProcessName = Left$(GetProcessName, InStr(GetProcessName, vbNullChar) - 1)
End Function

Public Sub DebugLog(sModule As String, sFunction As String, sText As String, Optional ByVal eType As LogEventTypeConstants = vbLogEventTypeInformation)
    Dim sPrefix         As String
    
    sPrefix = Format$(Now, FORMAT_TIME_ONLY) & Right$(Format$(Timer, FORMAT_BASE_3), 4) & ": "
    Select Case eType
    Case vbLogEventTypeError
        sPrefix = sPrefix & "[" & MSG_PREFIX_ERROR & "] "
    Case vbLogEventTypeWarning
        sPrefix = sPrefix & "[" & MSG_PREFIX_WARNING & "] "
    Case vbLogEventTypeDebug
        sPrefix = sPrefix & "[" & MSG_PREFIX_DEBUG & "] "
    End Select
    sPrefix = sPrefix & IIf(Len(sText) > 200, Left$(sText, 200) & "...", sText)
    If eType = vbLogEventTypeError Then
        sPrefix = sPrefix & " [" & sModule & "." & sFunction & "]"
    End If
    If eType = vbLogEventTypeError Then
        ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sPrefix & vbCrLf
    Else
        ConsolePrint sPrefix & vbCrLf
    End If
End Sub
