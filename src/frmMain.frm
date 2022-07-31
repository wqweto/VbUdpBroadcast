VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Const MODULE_NAME As String = "frmMain"

'=========================================================================
' API
'=========================================================================

Private Const MIB_IPROUTE_TYPE_DIRECT As Long = 3

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function ws_inet_ntoa Lib "ws2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long

Private Type IP_HDR
    ip_verlen       As Byte         ' 4-bit IPv4 version
                                    ' 4-bit header length (in 32-bit words)
    ip_tos          As Byte         ' IP type of service
    ip_totallength  As Integer      ' Total length
    ip_id           As Integer      ' Unique identifier
    ip_offset       As Integer      ' Fragment offset field
    ip_ttl          As Byte         ' Time to live
    ip_protocol     As Byte         ' Protocol(TCP,UDP etc)
    ip_checksum     As Integer      ' IP checksum
    ip_srcaddr      As Long         ' Source address
    ip_destaddr     As Long         ' Dest address
End Type

Private Type UDP_HDR
    udp_srcport     As Integer      ' Source port no.
    udp_destport    As Integer      ' Dest. port no.
    udp_length      As Integer      ' Udp packet length
    udp_checksum    As Integer      ' Udp checksum
End Type

Private Type EXT_UDP_HDR
    ip              As IP_HDR
    udp             As UDP_HDR
End Type
Private Const sizeof_EXT_UDP_HDR As Long = 28

'=========================================================================
' Constants and variables
'=========================================================================

Private Const STR_BROADCAST             As String = "255.255.255.255"
Private Const STR_LOOPBACK              As String = "127.0.0.1"

Private WithEvents m_oRawSocket     As cAsyncSocket
Attribute m_oRawSocket.VB_VarHelpID = -1
Private m_vForwardTable             As Variant

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
' Methods
'=========================================================================

Public Function Init() As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vElem           As Variant
    
    Set m_oRawSocket = New cAsyncSocket
    If Not m_oRawSocket.Create(SocketType:=ucsSckRaw, SocketAddress:=STR_LOOPBACK, SocketProtocol:=ucsScpUDP) Then
        GoTo QH
    End If
    If Not m_oRawSocket.GetLocalHost(vbNullString, vbNullString, ForwardTable:=m_vForwardTable) Then
        GoTo QH
    End If
    If Verbose Then
        For Each vElem In m_vForwardTable
            DebugLog MODULE_NAME, FUNC_NAME, "ForwardEntry=" & Join(vElem, ", ")
        Next
    End If
    Load Me
    '--- success
    Init = True
QH:
End Function

Public Sub Terminate()
    Set m_oRawSocket = Nothing
    Unload Me
End Sub

Private Sub pvOnReceive(baBuffer() As Byte)
    Const FUNC_NAME     As String = "pvOnReceive"
    Const IDX_DEST      As Long = 0
    Const IDX_MASK      As Long = 1
    Const IDX_NEXTHOP   As Long = 2
    Const IDX_TYPE      As Long = 4
    Dim uHdr            As EXT_UDP_HDR
    Dim sSrcAddr        As String
    Dim vElem           As Variant
    Dim sNextHop        As String
    Dim bNeedRelay      As Boolean
    
    On Error GoTo EH
    If UBound(baBuffer) + 1 < sizeof_EXT_UDP_HDR Then
        GoTo QH
    End If
    Call CopyMemory(uHdr, baBuffer(0), sizeof_EXT_UDP_HDR)
    If uHdr.ip.ip_ttl <= 1 Then
        GoTo QH
    End If
    If pvFromSinAddr(uHdr.ip.ip_destaddr) <> STR_BROADCAST Then
        GoTo QH
    End If
    sSrcAddr = pvFromSinAddr(uHdr.ip.ip_srcaddr)
    For Each vElem In m_vForwardTable
        If vElem(IDX_DEST) = STR_BROADCAST And vElem(IDX_MASK) = STR_BROADCAST And vElem(IDX_TYPE) = MIB_IPROUTE_TYPE_DIRECT Then
            If vElem(IDX_NEXTHOP) = sSrcAddr Then
                bNeedRelay = True
                Exit For
            End If
        End If
    Next
    If bNeedRelay Then
        For Each vElem In m_vForwardTable
            If vElem(IDX_DEST) = STR_BROADCAST And vElem(IDX_MASK) = STR_BROADCAST And vElem(IDX_TYPE) = MIB_IPROUTE_TYPE_DIRECT Then
                sNextHop = vElem(IDX_NEXTHOP)
                If sNextHop <> sSrcAddr And sNextHop <> STR_LOOPBACK Then
                    If Verbose Then
                        DebugLog MODULE_NAME, FUNC_NAME, "sSrcAddr=" & sSrcAddr & ", sNextHop=" & sNextHop & ", lSrcPort=" & (uHdr.udp.udp_srcport And &HFFFF&) & ", lDestPort=" & (uHdr.udp.udp_destport And &HFFFF&) & ", ForwardEntry=" & Join(vElem, ", ")
                    End If
                    pvSendBroadcast sNextHop, uHdr.udp.udp_srcport And &HFFFF&, uHdr.udp.udp_destport And &HFFFF&, _
                        VarPtr(baBuffer(sizeof_EXT_UDP_HDR)), UBound(baBuffer) + 1 - sizeof_EXT_UDP_HDR
                End If
            End If
        Next
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvSendBroadcast(sSrcAddr As String, ByVal lSrcPort As Long, ByVal lDstPort As Long, ByVal lPtr As Long, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvSendBroadcast"
    
    With New cAsyncSocket
        If Not .Create(SocketType:=ucsSckDatagram, SocketAddress:=sSrcAddr, SocketPort:=lSrcPort) Then
            DebugLog MODULE_NAME, FUNC_NAME, .LastError.Description, vbLogEventTypeError
            GoTo QH
        End If
        .SockOpt(ucsSsoBroadcast) = 1
        .SockOpt(ucsSsoDontRoute) = 1
        .SockOpt(ucsSsoIpTimeToLive, ucsSolIP) = 1
        If .Send(lPtr, lSize, STR_BROADCAST, lDstPort) <> lSize Then
            DebugLog MODULE_NAME, FUNC_NAME, .LastError.Description, vbLogEventTypeError
            GoTo QH
        End If
    End With
QH:
End Sub

Private Function pvFromSinAddr(ByVal sin_addr As Long) As String
    pvFromSinAddr = pvToString(ws_inet_ntoa(sin_addr))
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlen(lPtr), 0)
        Call CopyMemory(ByVal pvToString, ByVal lPtr, Len(pvToString))
    End If
End Function

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oRawSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oRawSocket_OnReceive"
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If Not m_oRawSocket.ReceiveArray(baBuffer) Then
        DebugLog MODULE_NAME, FUNC_NAME, m_oRawSocket.LastError.Description, vbLogEventTypeError
        GoTo QH
    End If
    pvOnReceive baBuffer
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
