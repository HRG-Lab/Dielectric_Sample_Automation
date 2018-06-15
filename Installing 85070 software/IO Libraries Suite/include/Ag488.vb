' --------------------------------------------------------------------------------
'  VB .NET Implementation of the  Agilent 488 API definitions.  Use this file in
'  Visual Basic .NET to program GPIB I/O interfaces that support the Agilent 488 
'  (or National Instruments NI-488.2) I/O API, such as Agilent 82357A USB/GPIB 
'  converter or Agilent 82350B PCI GPIB card.  See the Agilent IO Libraries Suite 
'  documentation for more information on programming with Agilent 488.
'
'  Copyright (C) 2005 Agilent Technologies, Inc.
' --------------------------------------------------------------------------------
'  Title   : ag488.vb
'  Date    : 07-13-2005
'  Purpose : Agilent 488 definitions for Visual Basic .NET
' -------------------------------------------------------------------------
Imports System
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization.Formatters.Binary



'namespace Agilent.TMFramework.Connectivity
Public Delegate Function GpibNotifyCallback_t(ByVal localUd As Integer, ByVal localIbsta As Integer, ByVal localIberr As Integer, ByVal localIbcntl As Integer, ByVal refData As IntPtr) As Integer

' <summary>
' This class contains the definitions for the Agilent 488 functions and constant values
' for use in Visual Basic .NET.
' </summary>

<Obsolete("The file has been deprecated. Please use the Kt488.vb in '\Program Files\Keysight\IO Libraries Suite\include'.")>
Public Class Ag488Wrap

#Region "Implementation Fields and definitions"
    Protected Shared m_globalsInitialized As Boolean = False
    Protected Shared m_statusVariables As IntPtr
    Protected Shared m_ibstaPtr As IntPtr
    Protected Shared m_iberrPtr As IntPtr
    Protected Shared m_ibcntPtr As IntPtr
    Protected Shared m_ibcntlPtr As IntPtr

    <DllImportAttribute("gpib-32.dll", EntryPoint:="RegisterGpibGlobalsForThread", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function RegisterGpibGlobalsForThread(ByVal Longibsta As IntPtr, ByVal Longiberr As IntPtr, ByVal Longibcnt As IntPtr, ByVal ibcntl As IntPtr) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="UnregisterGpibGlobalsForThread", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function UnregisterGpibGlobalsForThread() As Integer
    End Function

    Private Sub Ag488Wrap()
    End Sub
#End Region

#Region "Implementation helper functions"
    Protected Shared Sub VerifyGpibGlobalsRegistered()
        If Not m_globalsInitialized Then
            m_statusVariables = Marshal.AllocHGlobal(16)
            m_ibstaPtr = m_statusVariables
            m_iberrPtr = New IntPtr(m_statusVariables.ToInt32() + 4)
            m_ibcntPtr = New IntPtr(m_statusVariables.ToInt32() + 8)
            m_ibcntlPtr = New IntPtr(m_statusVariables.ToInt32() + 12)

            Dim rc As Integer = RegisterGpibGlobalsForThread( _
             m_ibstaPtr, m_iberrPtr, m_ibcntPtr, m_ibcntlPtr)
            If rc = 0 Then
                m_globalsInitialized = True
            ElseIf rc = 1 Then
                rc = UnregisterGpibGlobalsForThread()
                rc = RegisterGpibGlobalsForThread( _
                 m_ibstaPtr, m_iberrPtr, m_ibcntPtr, m_ibcntlPtr)
                m_globalsInitialized = True
            ElseIf (rc = 2 Or rc = 3) Then
                rc = UnregisterGpibGlobalsForThread()
                Marshal.FreeHGlobal(m_statusVariables)
            Else
                Marshal.FreeHGlobal(m_statusVariables)
            End If
        End If
    End Sub
    Protected Shared Sub VerifyGpibGlobalsNotRegistered()
        If m_globalsInitialized Then
            UnregisterGpibGlobalsForThread()
            Marshal.FreeHGlobal(m_statusVariables)
            m_ibstaPtr = IntPtr.Zero
            m_iberrPtr = IntPtr.Zero
            m_ibcntPtr = IntPtr.Zero
            m_ibcntlPtr = IntPtr.Zero
            m_statusVariables = IntPtr.Zero
            m_globalsInitialized = False
        End If
    End Sub

#End Region

#Region "488.2 Import Implementation"
    '  Agilent 488 Function Prototypes  
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibfindA", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibfindA32(ByVal udname As String) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibbnaA", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibbnaA32(ByVal ud As Integer, ByVal udname As String) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrdfA", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrdfA32(ByVal ud As Integer, ByVal filename As String) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrtfA", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrtfA32(ByVal ud As Integer, ByVal filename As String) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibfindW", ExactSpelling:=True, CharSet:=CharSet.Unicode, SetLastError:=False)> _
    Protected Shared Function ibfindW32(ByVal udname As String) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibbnaW", ExactSpelling:=True, CharSet:=CharSet.Unicode, SetLastError:=False)> _
    Protected Shared Function ibbnaW32(ByVal ud As Integer, ByVal udname As String) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrdfW", ExactSpelling:=True, CharSet:=CharSet.Unicode, SetLastError:=False)> _
    Protected Shared Function ibrdfW32(ByVal ud As Integer, ByVal filename As String) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrtfW", ExactSpelling:=True, CharSet:=CharSet.Unicode, SetLastError:=False)> _
    Protected Shared Function ibwrtfW32(ByVal ud As Integer, ByVal filename As String) As Integer
    End Function

    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibask", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibask32(ByVal ud As Integer, ByVal optionVal As Integer, <Out()> ByRef v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcac", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcac32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibclr", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibclr32(ByVal ud As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcmd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcmd32(ByVal ud As Integer, ByVal buf() As Byte, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcmd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcmd32(ByVal ud As Integer, ByVal buf() As Short, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcmd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcmd32(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcmd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcmd32(ByVal ud As Integer, ByVal buf() As Single, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcmd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcmd32(ByVal ud As Integer, ByVal buf() As Double, ByVal cnt As Integer) As Integer
    End Function

    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcmd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcmd32(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcmd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcmd32(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcmda", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcmda32(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibcmda", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibcmda32(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibconfig", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibconfig32(ByVal ud As Integer, ByVal optionVal As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibdev", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibdev32(ByVal boardID As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibdma", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibdma32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibeos", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibeos32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibeot", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibeot32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibgts", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibgts32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibist", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibist32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="iblck", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function iblck32(ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Integer, ByVal Reserved As IntPtr) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="iblines", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function iblines32(ByVal ud As Integer, <Out()> ByRef result As Short) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibln", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibln32(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, <Out()> ByRef listen As Short) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibloc", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibloc32(ByVal ud As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibnotify", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibnotify32(ByVal ud As Integer, ByVal mask As Integer, ByVal Callback As GpibNotifyCallback_t, ByVal RefData As IntPtr) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibonl", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibonl32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibpad", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibpad32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibpct", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibpct32(ByVal ud As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibpoke", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibpoke32(ByVal ud As Integer, ByVal optionVal As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibppc", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibppc32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrd32(ByVal ud As Integer, ByVal buf() As Byte, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrd32(ByVal ud As Integer, ByVal buf() As Short, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrd32(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrd32(ByVal ud As Integer, ByVal buf() As Single, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrd32(ByVal ud As Integer, ByVal buf() As Double, ByVal cnt As Integer) As Integer
    End Function

    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrd32(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrd", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrd32(ByVal ud As Integer, ByVal buf As StringBuilder, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrda", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrda32(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrda", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrda32(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrpp", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrpp32(ByVal ud As Integer, <Out()> ByRef ppr As Byte) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrsc", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrsc32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrsp", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrsp32(ByVal ud As Integer, <Out()> ByRef spr As Byte) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibrsv", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibrsv32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibsad", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibsad32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibsic", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibsic32(ByVal ud As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibsre", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibsre32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibstop", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibstop32(ByVal ud As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibtmo", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibtmo32(ByVal ud As Integer, ByVal v As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibtrg", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibtrg32(ByVal ud As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwait", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwait32(ByVal ud As Integer, ByVal mask As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrt", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrt32(ByVal ud As Integer, ByVal buf() As Byte, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrt", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrt32(ByVal ud As Integer, ByVal buf() As Short, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrt", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrt32(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrt", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrt32(ByVal ud As Integer, ByVal buf() As Single, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrt", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrt32(ByVal ud As Integer, ByVal buf() As Double, ByVal cnt As Integer) As Integer
    End Function

    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrt", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrt32(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrt", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrt32(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrta", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrta32(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrta", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrta32(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibwrta", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibwrta32(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
    End Function

    ' GPIB-ENET only functions to support locking across machines
    ' Deprecated - Use iblck
    <DllImportAttribute("gpib-32.dll", EntryPoint:="iblock", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function iblock32(ByVal ud As Integer) As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ibunlock", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ibunlock32(ByVal ud As Integer) As Integer
    End Function

    '************************************************************************
    '  Functions to access Thread-Specific copies of the GPIB global vars 

    <DllImportAttribute("gpib-32.dll", EntryPoint:="ThreadIbsta", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ThreadIbsta32() As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ThreadIberr", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ThreadIberr32() As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ThreadIbcnt", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ThreadIbcnt32() As Integer
    End Function
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ThreadIbcntl", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Function ThreadIbcntl32() As Integer
    End Function


    '************************************************************************
    '  Agilent 488 Function Prototypes  

    <DllImportAttribute("gpib-32.dll", EntryPoint:="AllSpoll", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub AllSpoll32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal results() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="DevClear", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub DevClear32(ByVal boardID As Integer, ByVal addr As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="DevClearList", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub DevClearList32(ByVal boardID As Integer, ByVal addrlist() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="EnableLocal", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub EnableLocal32(ByVal boardID As Integer, ByVal addrlist() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="EnableRemote", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub EnableRemote32(ByVal boardID As Integer, ByVal addrlist() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="FindLstn", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub FindLstn32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal results() As Short, ByVal limit As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="FindRQS", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub FindRQS32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal dev_stat() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="PPoll", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub PPoll32(ByVal boardID As Integer, <Out()> ByRef result As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="PPollConfig", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub PPollConfig32(ByVal boardID As Integer, ByVal addr As Short, ByVal dataLine As Integer, ByVal lineSense As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="PPollUnconfig", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub PPollUnconfig32(ByVal boardID As Integer, ByVal addrlist() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="PassControl", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub PassControl32(ByVal boardID As Integer, ByVal addr As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="RcvRespMsg", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub RcvRespMsg32(ByVal boardID As Integer, ByVal buffer() As Byte, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="RcvRespMsg", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub RcvRespMsg32(ByVal boardID As Integer, ByVal buffer() As Short, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="RcvRespMsg", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub RcvRespMsg32(ByVal boardID As Integer, ByVal buffer() As Integer, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="RcvRespMsg", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub RcvRespMsg32(ByVal boardID As Integer, ByVal buffer() As Single, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="RcvRespMsg", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub RcvRespMsg32(ByVal boardID As Integer, ByVal buffer() As Double, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub

    <DllImportAttribute("gpib-32.dll", EntryPoint:="RcvRespMsg", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub RcvRespMsg32(ByVal boardID As Integer, ByVal buffer As IntPtr, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="RcvRespMsg", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub RcvRespMsg32(ByVal boardID As Integer, ByVal buffer As StringBuilder, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ReadStatusByte", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub ReadStatusByte32(ByVal boardID As Integer, ByVal addr As Short, <Out()> ByRef result As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Receive", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Receive32(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Byte, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Receive", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Receive32(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Short, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Receive", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Receive32(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Integer, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Receive", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Receive32(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Single, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Receive", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Receive32(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Double, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub

    <DllImportAttribute("gpib-32.dll", EntryPoint:="Receive", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Receive32(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer As StringBuilder, ByVal cnt As Integer, ByVal Termination As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ReceiveSetup", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub ReceiveSetup32(ByVal boardID As Integer, ByVal addr As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="ResetSys", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub ResetSys32(ByVal boardID As Integer, ByVal addrlist() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Send", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Send32(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Byte, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Send", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Send32(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Short, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Send", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Send32(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Integer, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Send", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Send32(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Single, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Send", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Send32(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Double, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub

    <DllImportAttribute("gpib-32.dll", EntryPoint:="Send", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Send32(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf As IntPtr, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Send", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Send32(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf As String, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendCmds", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendCmds32(ByVal boardID As Integer, ByVal buffer() As Byte, ByVal cnt As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendCmds", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendCmds32(ByVal boardID As Integer, ByVal buffer() As Short, ByVal cnt As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendCmds", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendCmds32(ByVal boardID As Integer, ByVal buffer() As Integer, ByVal cnt As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendCmds", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendCmds32(ByVal boardID As Integer, ByVal buffer() As Single, ByVal cnt As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendCmds", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendCmds32(ByVal boardID As Integer, ByVal buffer() As Double, ByVal cnt As Integer)
    End Sub

    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendCmds", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendCmds32(ByVal boardID As Integer, ByVal buffer As IntPtr, ByVal cnt As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendCmds", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendCmds32(ByVal boardID As Integer, ByVal buffer As String, ByVal cnt As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendDataBytes", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendDataBytes32(ByVal boardID As Integer, ByVal buffer() As Byte, ByVal cnt As Integer, ByVal eot_mode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendDataBytes", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendDataBytes32(ByVal boardID As Integer, ByVal buffer() As Short, ByVal cnt As Integer, ByVal eot_mode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendDataBytes", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendDataBytes32(ByVal boardID As Integer, ByVal buffer() As Integer, ByVal cnt As Integer, ByVal eot_mode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendDataBytes", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendDataBytes32(ByVal boardID As Integer, ByVal buffer() As Single, ByVal cnt As Integer, ByVal eot_mode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendDataBytes", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendDataBytes32(ByVal boardID As Integer, ByVal buffer() As Double, ByVal cnt As Integer, ByVal eot_mode As Integer)
    End Sub

    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendDataBytes", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendDataBytes32(ByVal boardID As Integer, ByVal buffer As IntPtr, ByVal cnt As Integer, ByVal eot_mode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendDataBytes", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendDataBytes32(ByVal boardID As Integer, ByVal buffer As String, ByVal cnt As Integer, ByVal eot_mode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendIFC", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendIFC32(ByVal boardID As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendLLO", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendLLO32(ByVal boardID As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendList", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendList32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Byte, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendList", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendList32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Short, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendList", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendList32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Integer, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendList", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendList32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Single, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendList", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendList32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Double, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub

    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendList", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendList32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf As IntPtr, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendList", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendList32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf As String, ByVal datacnt As Integer, ByVal eotMode As Integer)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SendSetup", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SendSetup32(ByVal boardID As Integer, ByVal addrlist() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="SetRWLS", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub SetRWLS32(ByVal boardID As Integer, ByVal addrlist() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="TestSRQ", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub TestSRQ32(ByVal boardID As Integer, <Out()> ByRef result As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="TestSys", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub TestSys32(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal results() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="Trigger", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub Trigger32(ByVal boardID As Integer, ByVal addr As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="TriggerList", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub TriggerList32(ByVal boardID As Integer, ByVal addrlist() As Short)
    End Sub
    <DllImportAttribute("gpib-32.dll", EntryPoint:="WaitSRQ", ExactSpelling:=True, CharSet:=CharSet.Ansi, SetLastError:=False)> _
    Protected Shared Sub WaitSRQ32(ByVal boardID As Integer, <Out()> ByRef result As Short)
    End Sub
#End Region

#Region "488.2 Constants"

    '*************************************************************************
    '    HANDY CONSTANTS FOR USE BY APPLICATION PROGRAMS ...                  
    '*************************************************************************
    Public Const UNL As Byte = &H3F   '   GPIB unlisten command                 
    Public Const UNT As Byte = &H5F   '   GPIB untalk command                   
    Public Const GTL As Byte = &H1    '   GPIB go to local                      
    Public Const SDC As Byte = &H4    '   GPIB selected device clear            
    Public Const PPC As Byte = &H5    '   GPIB parallel poll configure          
    Public Const GGET As Byte = &H8    '   GPIB group execute trigger            
    Public Const TCT As Byte = &H9    '   GPIB take control                     
    Public Const LLO As Byte = &H11   '   GPIB local lock out                   
    Public Const DCL As Byte = &H14   '   GPIB device clear                     
    Public Const PPU As Byte = &H15   '   GPIB parallel poll unconfigure        
    Public Const SPE As Byte = &H18   '   GPIB serial poll enable               
    Public Const SPD As Byte = &H19   '   GPIB serial poll disable              
    Public Const PPE As Byte = &H60   '   GPIB parallel poll enable             
    Public Const PPD As Byte = &H70   '   GPIB parallel poll disable            

    ' GPIB status bit vector :                                 
    '       global variable ibsta and wait mask                

    Public Const EERR As Integer = (1 << 15) '   Error detected                   		
    Public Const ERR As Integer = (1 << 15) '   Error detected                   
    Public Const TIMO As Integer = (1 << 14) '   Timeout                          
    Public Const EEND As Integer = (1 << 13) '   EOI or EOS detected              
    Public Const [END] As Integer = (1 << 13) '   EOI or EOS detected              
    Public Const SRQI As Integer = (1 << 12) '   SRQ detected by CIC              
    Public Const RQS As Integer = (1 << 11) '   Device needs service             
    Public Const CMPL As Integer = (1 << 8) '   I/O completed                    
    Public Const LOK As Integer = (1 << 7) '   Local lockout state              
    Public Const RREM As Integer = (1 << 6) '   Remote state                     
    Public Const CIC As Integer = (1 << 5) '   Controller-in-Charge             
    Public Const ATN As Integer = (1 << 4) '   Attention asserted               
    Public Const TACS As Integer = (1 << 3) '   Talker active                    
    Public Const LACS As Integer = (1 << 2) '   Listener active                  
    Public Const DTAS As Integer = (1 << 1) '   Device trigger state             
    Public Const DCAS As Integer = (1 << 0) '   Device clear state               

    ' Error messages returned in global variable iberr         

    Public Const EDVR As Integer = 0   '   System error                            
    Public Const ECIC As Integer = 1   '   Function requires GPIB board to be CIC  
    Public Const ENOL As Integer = 2   '   Write function detected no Listeners    
    Public Const EADR As Integer = 3   '   Interface board not addressed correctly 
    Public Const EARG As Integer = 4   '   Invalid argument to function call       
    Public Const ESAC As Integer = 5   '   Function requires GPIB board to be SAC  
    Public Const EABO As Integer = 6   '   I/O operation aborted                   
    Public Const ENEB As Integer = 7   '   Non-existent interface board            
    Public Const EDMA As Integer = 8   '   Error performing DMA                    
    Public Const EOIP As Integer = 10   '   I/O operation started before previous   
    ' operation completed                     
    Public Const ECAP As Integer = 11   '   No capability for intended operation    
    Public Const EFSO As Integer = 12   '   File system operation error             
    Public Const EBUS As Integer = 14   '   Command error during device call        
    Public Const ESTB As Integer = 15   '   Serial poll status byte lost            
    Public Const ESRQ As Integer = 16   '   SRQ remains asserted                    
    Public Const ETAB As Integer = 20   '   The return buffer is full.              
    Public Const ELCK As Integer = 21   '   Address or board is locked.             
    Public Const EARM As Integer = 22   '   The ibnotify Callback failed to rearm   
    Public Const EHDL As Integer = 23   '   The input handle is invalid             
    Public Const EWIP As Integer = 26   '   Wait already in progress on input ud    
    Public Const ERST As Integer = 27   '   The event notification was cancelled    
    ' due to a reset of the interface         
    Public Const EPWR As Integer = 28   '   The system or board has lost power or   
    ' gone to standby                         

    ' Warning messages returned in global variable iberr       

    Public Const WCFG As Integer = 24   '   Configuration warning                   
    Public Const ECFG As Integer = WCFG

    ' EOS mode bits                                            

    Public Const BIN As Integer = (1 << 12) '   Eight bit compare                   
    Public Const XEOS As Integer = (1 << 11) '   Send END with EOS byte              
    Public Const REOS As Integer = (1 << 10) '   Terminate read on EOS               

    ' Timeout values and meanings                              

    Public Const TNONE As Integer = 0    '   Infinite timeout (disabled)         
    Public Const T10us As Integer = 1    '   Timeout of 10 us (ideal)            
    Public Const T30us As Integer = 2    '   Timeout of 30 us (ideal)            
    Public Const T100us As Integer = 3    '   Timeout of 100 us (ideal)           
    Public Const T300us As Integer = 4    '   Timeout of 300 us (ideal)           
    Public Const T1ms As Integer = 5    '   Timeout of 1 ms (ideal)             
    Public Const T3ms As Integer = 6    '   Timeout of 3 ms (ideal)             
    Public Const T10ms As Integer = 7    '   Timeout of 10 ms (ideal)            
    Public Const T30ms As Integer = 8    '   Timeout of 30 ms (ideal)            
    Public Const T100ms As Integer = 9    '   Timeout of 100 ms (ideal)           
    Public Const T300ms As Integer = 10    '   Timeout of 300 ms (ideal)           
    Public Const T1s As Integer = 11    '   Timeout of 1 s (ideal)              
    Public Const T3s As Integer = 12    '   Timeout of 3 s (ideal)              
    Public Const T10s As Integer = 13    '   Timeout of 10 s (ideal)             
    Public Const T30s As Integer = 14    '   Timeout of 30 s (ideal)             
    Public Const T100s As Integer = 15    '   Timeout of 100 s (ideal)            
    Public Const T300s As Integer = 16    '   Timeout of 300 s (ideal)            
    Public Const T1000s As Integer = 17    '   Timeout of 1000 s (ideal)           

    '  IBLN Constants                                          
    Public Const NO_SAD As Integer = 0 ' 
    Public Const ALL_SAD As Integer = -1

    '  The following constants are used for the second parameter of the
    '  ibconfig function.  They are the "option" selection codes.

    Public Const IbcPAD As Integer = &H1          '   Primary Address                      
    Public Const IbcSAD As Integer = &H2          '   Secondary Address                    
    Public Const IbcTMO As Integer = &H3          '   Timeout Value                        
    Public Const IbcEOT As Integer = &H4          '   Send EOI with last data byte?        
    Public Const IbcPPC As Integer = &H5          '   Parallel Poll Configure              
    Public Const IbcREADDR As Integer = &H6          '   Repeat Addressing                    
    Public Const IbcAUTOPOLL As Integer = &H7          '   Disable Auto Serial Polling          
    Public Const IbcCICPROT As Integer = &H8          '   Use the CIC Protocol?                
    Public Const IbcIRQ As Integer = &H9          '   Use PIO for I/O                      
    Public Const IbcSC As Integer = &HA          '   Board is System Controller?          
    Public Const IbcSRE As Integer = &HB          '   Assert SRE on device calls?          
    Public Const IbcEOSrd As Integer = &HC          '   Terminate reads on EOS               
    Public Const IbcEOSwrt As Integer = &HD          '   Send EOI with EOS character          
    Public Const IbcEOScmp As Integer = &HE          '   Use 7 or 8-bit EOS compare           
    Public Const IbcEOSchar As Integer = &HF          '   The EOS character.                   
    Public Const IbcPP2 As Integer = &H10         '   Use Parallel Poll Mode 2.            
    Public Const IbcTIMING As Integer = &H11         '   NORMAL, HIGH, or VERY_HIGH timing.   
    Public Const IbcDMA As Integer = &H12         '   Use DMA for I/O                      
    Public Const IbcReadAdjust As Integer = &H13         '   Swap bytes during an ibrd.           
    Public Const IbcWriteAdjust As Integer = &H14        '   Swap bytes during an ibwrt.          
    Public Const IbcSendLLO As Integer = &H17         '   Enable/disable the sending of LLO.      
    Public Const IbcSPollTime As Integer = &H18         '   Set the timeout value for serial polls. 
    Public Const IbcPPollTime As Integer = &H19         '   Set the parallel poll length period.    
    Public Const IbcEndBitIsNormal As Integer = &H1A     '   Remove EOS from END bit of IBSTA.       
    Public Const IbcUnAddr As Integer = &H1B     '   Enable/disable device unaddressing.     
    Public Const IbcSignalNumber As Integer = &H1C     '   Set UNIX signal number - unsupported 
    Public Const IbcBlockIfLocked As Integer = &H1D     '   Enable/disable blocking for locked boards/devices 
    Public Const IbcHSCableLength As Integer = &H1F     '   Length of cable specified for high speed timing.
    Public Const IbcIst As Integer = &H20         '   Set the IST bit.                     
    Public Const IbcRsv As Integer = &H21         '   Set the RSV byte.                    
    Public Const IbcLON As Integer = &H22         '   Enter listen only mode               

    '
    '    Constants that can be used (ByVal addition as in to the ibconfig constants)
    '    when calling the ibask(ByVal ) as function.


    Public Const IbaPAD As Integer = IbcPAD
    Public Const IbaSAD As Integer = IbcSAD
    Public Const IbaTMO As Integer = IbcTMO
    Public Const IbaEOT As Integer = IbcEOT
    Public Const IbaPPC As Integer = IbcPPC
    Public Const IbaREADDR As Integer = IbcREADDR
    Public Const IbaAUTOPOLL As Integer = IbcAUTOPOLL
    Public Const IbaCICPROT As Integer = IbcCICPROT
    Public Const IbaIRQ As Integer = IbcIRQ
    Public Const IbaSC As Integer = IbcSC
    Public Const IbaSRE As Integer = IbcSRE
    Public Const IbaEOSrd As Integer = IbcEOSrd
    Public Const IbaEOSwrt As Integer = IbcEOSwrt
    Public Const IbaEOScmp As Integer = IbcEOScmp
    Public Const IbaEOSchar As Integer = IbcEOSchar
    Public Const IbaPP2 As Integer = IbcPP2
    Public Const IbaTIMING As Integer = IbcTIMING
    Public Const IbaDMA As Integer = IbcDMA
    Public Const IbaReadAdjust As Integer = IbcReadAdjust
    Public Const IbaWriteAdjust As Integer = IbcWriteAdjust
    Public Const IbaSendLLO As Integer = IbcSendLLO
    Public Const IbaSPollTime As Integer = IbcSPollTime
    Public Const IbaPPollTime As Integer = IbcPPollTime
    Public Const IbaEndBitIsNormal As Integer = IbcEndBitIsNormal
    Public Const IbaUnAddr As Integer = IbcUnAddr
    Public Const IbaSignalNumber As Integer = IbcSignalNumber
    Public Const IbaBlockIfLocked As Integer = IbcBlockIfLocked
    Public Const IbaHSCableLength As Integer = IbcHSCableLength
    Public Const IbaIst As Integer = IbcIst
    Public Const IbaRsv As Integer = IbcRsv
    Public Const IbaLON As Integer = IbcLON
    Public Const IbaSerialNumber As Integer = &H23

    Public Const IbaBNA As Integer = &H200     '   A device's access board. 


    ' Values used by the Send 488.2 command. 
    Public Const NULLend As Integer = &H0    '   Do nothing at the end of a transfer.
    Public Const NLend As Integer = &H1    '   Send NL with EOI after a transfer.  
    Public Const DABend As Integer = &H2    '   Send EOI with the last DAB.         

    ' Value used by the 488.2 Receive command.

    Public Const STOPend As Integer = &H100

    '
    '  This value is used to terminate an address list.  It should be
    '  assigned to the last entry.


    Public Const NOADDR As Short = &HFFFFS



    ' iblines constants 

    Public Const ValidEOI As Short = &H80
    Public Const ValidATN As Short = &H40
    Public Const ValidSRQ As Short = &H20
    Public Const ValidREN As Short = &H10
    Public Const ValidIFC As Short = &H8
    Public Const ValidNRFD As Short = &H4
    Public Const ValidNDAC As Short = &H2
    Public Const ValidDAV As Short = &H1
    Public Const BusEOI As Short = &H8000S
    Public Const BusATN As Short = &H4000
    Public Const BusSRQ As Short = &H2000
    Public Const BusREN As Short = &H1000
    Public Const BusIFC As Short = &H800
    Public Const BusNRFD As Short = &H400
    Public Const BusNDAC As Short = &H200
    Public Const BusDAV As Short = &H100
#End Region

#Region "488.2 Status Variables"

    Public Shared Sub RegisterGpibGlobals()
        VerifyGpibGlobalsRegistered()
    End Sub

    Public Shared Sub UnregisterGPIBGlobals()
        VerifyGpibGlobalsNotRegistered()
    End Sub

    Public Shared Property ibsta() As Integer
        Get
            VerifyGpibGlobalsRegistered()
            If m_globalsInitialized Then
                Return Marshal.ReadInt32(m_ibstaPtr)
            Else
                Return &H8000
            End If
        End Get
        Set(ByVal value As Integer)
            VerifyGpibGlobalsRegistered()
            If m_globalsInitialized Then
                Marshal.WriteInt32(m_ibstaPtr, value)
            End If
        End Set
    End Property

    Public Shared Property iberr() As Integer
        Get
            VerifyGpibGlobalsRegistered()
            If m_globalsInitialized Then
                Return Marshal.ReadInt32(m_iberrPtr)
            Else
                Return EDVR
            End If
        End Get
        Set(ByVal value As Integer)
            VerifyGpibGlobalsRegistered()
            If m_globalsInitialized Then
                Marshal.WriteInt32(m_iberrPtr, value)
            End If
        End Set
    End Property

    Public Shared Property ibcnt() As Integer
        Get
            VerifyGpibGlobalsRegistered()
            If m_globalsInitialized Then
                Return Marshal.ReadInt32(m_ibcntPtr)
            Else
                Return EDVR
            End If
        End Get
        Set(ByVal value As Integer)
            VerifyGpibGlobalsRegistered()
            If m_globalsInitialized Then
                Marshal.WriteInt32(m_ibcntPtr, value)
            End If
        End Set
    End Property

    Public Shared Property ibcntl() As Integer
        Get
            VerifyGpibGlobalsRegistered()
            If m_globalsInitialized Then
                Return Marshal.ReadInt32(m_ibcntlPtr)
            Else
                Return EDVR
            End If
        End Get
        Set(ByVal value As Integer)
            VerifyGpibGlobalsRegistered()
            If m_globalsInitialized Then
                Marshal.WriteInt32(m_ibcntlPtr, value)
            End If
        End Set
    End Property

#End Region

#Region "488.2 Public Methods"
    '  Agilent 488 Function Prototypes  

    ''' <summary>
    ''' Find a device or board descriptor by name
    ''' </summary>
    ''' <param name="udname">Name to find</param>
    ''' <returns>Device or board descriptor</returns>
    Public Shared Function ibfind(ByVal udname As String) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibfindW32(udname)
    End Function

    Public Shared Function ibbna(ByVal ud As Integer, ByVal udname As String) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibbnaW32(ud, udname)
    End Function

    ''' <summary>
    ''' Transfer data from GPIB to a file
    ''' </summary>
    ''' <param name="ud">Device or board descriptor of source</param>
    ''' <param name="filename">File name to receive data</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibrdf(ByVal ud As Integer, ByVal filename As String) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrdfW32(ud, filename)
    End Function

    ''' <summary>
    ''' Write data from a file to the GPIB
    ''' </summary>
    ''' <param name="ud">Device or board descriptor of destination</param>
    ''' <param name="filename">File name of the data source</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibwrtf(ByVal ud As Integer, ByVal filename As String) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrtfW32(ud, filename)
    End Function

    ''' <summary>
    ''' Return the value of a configuration parameter
    ''' </summary>
    ''' <param name="ud">Device or board descriptor</param>
    ''' <param name="option">Parameter to return (IbaXXX)</param>
    ''' <param name="v">Output variable to receive value</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibask(ByVal ud As Integer, ByVal optionVal As Integer, <Out()> ByRef v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibask32(ud, optionVal, v)
    End Function

    ''' <summary>
    ''' Make the interface Active Controller
    ''' </summary>
    ''' <param name="ud">Board descriptor to make active</param>
    ''' <param name="v">0 for asynchronous operation, anything else for synchronous</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>Call ibsic before ibcac to ensure this board is Controller in Charge</remarks>
    Public Shared Function ibcac(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcac32(ud, v)
    End Function

    ''' <summary>
    ''' Clear a device
    ''' </summary>
    ''' <param name="ud">Device descriptor to clear</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibclr(ByVal ud As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibclr32(ud)
    End Function

    ' Note for Function ibcmd(ByVal ud as Integer, ByVal buf() as byte, ByVal cnt as Integer) as Integer
    '	This Method's "buf" parameter is passed directly to the API
    '	implementation. 
    ''' <summary>
    ''' Send a GPIB command
    ''' </summary>
    ''' <param name="ud">Board descriptor to send the command</param>
    ''' <param name="buf">Buffer to send </param>
    ''' <param name="cnt">Number of bytes to send from 'buf'</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibcmd(ByVal ud As Integer, ByVal buf() As Byte, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcmd32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Send a GPIB command
    ''' </summary>
    ''' <param name="ud">Board descriptor to send the command</param>
    ''' <param name="buf">Buffer to send</param>
    ''' <param name="cnt">Number of bytes to send from 'buf'</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibcmd(ByVal ud As Integer, ByVal buf() As Short, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcmd32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Send a GPIB command
    ''' </summary>
    ''' <param name="ud">Board descriptor to send the command</param>
    ''' <param name="buf">Buffer to send</param>
    ''' <param name="cnt">Number of bytes to send from 'buf'</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibcmd(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcmd32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Send a GPIB command
    ''' </summary>
    ''' <param name="ud">Board descriptor to send the command</param>
    ''' <param name="buf">Buffer to send</param>
    ''' <param name="cnt">Number of bytes to send from 'buf'</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibcmd(ByVal ud As Integer, ByVal buf() As Single, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcmd32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Send a GPIB command
    ''' </summary>
    ''' <param name="ud">Board descriptor to send the command</param>
    ''' <param name="buf">Buffer to send</param>
    ''' <param name="cnt">Number of bytes to send from 'buf'</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibcmd(ByVal ud As Integer, ByVal buf() As Double, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcmd32(ud, buf, cnt)
    End Function

    ' Note for Function ibcmd(ByVal ud as Integer, ByVal buf as IntPtr, ByVal cnt as Integer) as Integer
    '	This Method's "buf" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "buf" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    ''' <summary>
    ''' Send a GPIB command
    ''' </summary>
    ''' <param name="ud">Board descriptor to send the command</param>
    ''' <param name="buf">Buffer to send</param>
    ''' <param name="cnt">Number of bytes to send from 'buf'</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' The 'buf' parameter is passed directly to .NET's unmanaged interop layer.
    ''' Use GCHandle.Alloc to get a pinned pointer to a byte array, and
    ''' pass handle.AddrOfPinnedObject() as 'buf'.
    ''' </remarks>
    Public Shared Function ibcmd(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcmd32(ud, buf, cnt)
    End Function

    ' Note for Function ibcmd(ByVal ud as Integer, ByVal buf as string, ByVal cnt as Integer) as Integer
    '	This Method's "buf" parameter is treated as an ASCII string that 
    '	is converted by .NET into an array of bytes.  Consider using 
    '	a different version of this method for passing binary data. 
    ''' <summary>
    ''' Send a GPIB command
    ''' </summary>
    ''' <param name="ud">Board descriptor to send the command</param>
    ''' <param name="buf">Buffer to send</param>
    ''' <param name="cnt">Number of bytes to send from 'buf'</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' 'buf' is treated here as an ASCII string, which .NET converts to an array of bytes.
    ''' Consider using a byte array for 'buf' if you want to pass binary data.
    Public Shared Function ibcmd(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcmd32(ud, buf, cnt)
    End Function

    ' Note: the buf parameter must point to a GCHandle to a pinned .NET
    ' * object (see GCHandle and .NET interoperation document for more
    ' * information) containing the data to be sent to the device 
    ''' <summary>
    ''' Send a GPIB command asyncronously
    ''' </summary>
    ''' <param name="ud">Board descriptor to send the command</param>
    ''' <param name="buf">Buffer to send</param>
    ''' <param name="cnt">Number of bytes to send from 'buf'</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' The 'buf' parameter is passed directly to .NET's unmanaged interop layer.
    ''' Use GCHandle.Alloc to get a pinned pointer to a byte array, and
    ''' pass handle.AddrOfPinnedObject() as 'buf'.
    ''' Use 'ibwait' to wait for completion of this operation.
    ''' </remarks>		
    Public Shared Function ibcmda(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcmda32(ud, buf, cnt)
    End Function

    ' Note for Function ibcmda(ByVal ud as Integer, ByVal buf as string, ByVal cnt as Integer) as Integer
    '	This Method's "buf" parameter is treated as an ASCII string that 
    '	is converted by .NET into an array of bytes.  Consider using 
    '	a different version of this method for passing binary data. 
    Public Shared Function ibcmda(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibcmda32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Set the value of a configuration parameter.
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="option">Parameter to set (IbcXXX)</param>
    ''' <param name="v">Value for the parameter</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibconfig(ByVal ud As Integer, ByVal optionVal As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibconfig32(ud, optionVal, v)
    End Function

    ''' <summary>
    ''' Get a device descriptor
    ''' </summary>
    ''' <param name="boardID">Board number to which the device is connected</param>
    ''' <param name="pad">Primary address of the device</param>
    ''' <param name="sad">Secondary address of the device (0 if none)</param>
    ''' <param name="tmo">Timeout for the device (Txxx constants: see ibtmo reference in help)</param>
    ''' <param name="eot">1 to enable end-of-transmission EOI, 0 to disable EOI</param>
    ''' <param name="eos">0 to disable end-of-string termination, nonzero to enable (see ibeos)</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibdev(ByVal boardID As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibdev32(boardID, pad, sad, tmo, eot, eos)
    End Function

    ''' <summary>
    ''' Enable or disable DMA (direct memory access)
    ''' </summary>
    ''' <param name="ud">Board (interface) descriptor</param>
    ''' <param name="v">0 to disable DMA, 1 to enable</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibdma(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibdma32(ud, v)
    End Function


    ''' <summary>
    ''' Enable or disable end-of-string termination
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="v">0 to disable EOS termination.  See Remarks for other values.</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' The 'v' parameter is a 16-bit value which contains a mode setting in the high byte
    ''' and a EOS character in the low byte.
    ''' Possible high byte values:
    '''		0x04	Terminate read on detecting EOS
    '''		0x08	Set EOT with EOS on write
    '''		0x10	Use 8-bit EOS comparison (the fixed default for Agilent interfaces)
    '''	The EOS byte is not send automatically on writes: include it at the end of written data strings. 
    ''' </remarks>
    Public Shared Function ibeos(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibeos32(ud, v)
    End Function

    ''' <summary>
    ''' Enable or disable assertion of EOI at the end of writes.
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="v">0 to disable EOI, nonzero to enable</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibeot(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibeot32(ud, v)
    End Function

    ''' <summary>
    ''' Make the interface Standby Controller
    ''' </summary>
    ''' <param name="ud">Board (interface) descriptor</param>
    ''' <param name="v">Nonzero to perform acceptor handshaking (ignored by Agilent interfaces)</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibgts(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibgts32(ud, v)
    End Function

    ''' <summary>
    ''' Set or clear ist (individual status) bit
    ''' </summary>
    ''' <param name="ud">Board (interface) descriptor</param>
    ''' <param name="v">Nonzero to set, zero to clear</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibist(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibist32(ud, v)
    End Function

    ' Note for Function iblck(ByVal ud as Integer, ByVal v as Integer, ByVal LockWaitTime as uInteger, ByVal Reserved as IntPtr) as Integer
    '	Pass IntPtr.Zero into this method's "Reserved" parameter. 
    ''' <summary>
    ''' Acquire or release an interface lock
    ''' </summary>
    ''' <param name="ud">Board (interface) descriptor</param>
    ''' <param name="v">Nonzero to acquire, zero to release</param>
    ''' <param name="LockWaitTime">Time in milliseconds to wait for a lock</param>
    ''' <param name="Reserved">Must be IntPtr.Zero</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function iblck(ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Integer, ByVal Reserved As IntPtr) As Integer
        VerifyGpibGlobalsRegistered()
        Return iblck32(ud, v, LockWaitTime, Reserved)
    End Function

    ''' <summary>
    ''' Return the state of the bus management lines
    ''' </summary>
    ''' <param name="ud">Board descriptor</param>
    ''' <param name="result">Bus management line status information (bit per line)</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' Status line bits:
    ''' Bit 7	EOI
    ''' Bit 6	ATN
    ''' Bit 5	SRQ
    ''' Bit 4	REN
    ''' Bit 3	IFC
    ''' Bit 2	NRFD
    ''' Bit 1	NDAC
    ''' Bit 0	DAV
    ''' The low byte (bits 7-0) indicates whether the interface can sense the line
    ''' The high byte (bits 15-8) indicates the line status for each line (same bit order)
    ''' </remarks>
    Public Shared Function iblines(ByVal ud As Integer, <Out()> ByRef result As Short) As Integer
        VerifyGpibGlobalsRegistered()
        Return iblines32(ud, result)
    End Function

    ''' <summary>
    ''' Test for listeners
    ''' </summary>
    ''' <param name="ud">Board (interface) descriptor to check</param>
    ''' <param name="pad">Primary device address to test</param>
    ''' <param name="sad">Secondary device address to test (or NO_SAD or ALL_SAD)</param>
    ''' <param name="listen">Output, set to nonzero if listener present</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibln(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, <Out()> ByRef listen As Short) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibln32(ud, pad, sad, listen)
    End Function

    ''' <summary>
    ''' Place a device in local mode
    ''' </summary>
    ''' <param name="ud">Device or board descriptor (Agilent interfaces support only device-level)</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibloc(ByVal ud As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibloc32(ud)
    End Function

    ' Note: the RefData parameter must point to a GCHandle to a pinned .NET
    ' * object (ByVal GCHandle as see and .NET interoperation document for more
    ' * information) containing the custom user data passed into the 
    ' * callback function. 
    ''' <summary>
    ''' Register a callback function to be called when events occur
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="mask">GPIB event mask: see remarks below</param>
    ''' <param name="Callback">Callback function to call on event</param>
    ''' <param name="RefData">Reference data for the callback. Must be pinned GCHandle: see remarks</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' GPIB event mask bit constants:
    '''		TIMO	(0x4000)
    '''		END		(0x2000)
    '''		SRQI	(0x1000)
    '''		RQS		(0x800)
    '''		CMPL	(0x100)
    '''		LOK		(0x80)
    '''		REM		(0x40)
    '''		CIC		(0x20)
    '''		ATN		(0x10)
    '''		TACS	(0x8)
    '''		LACS	(0x4)
    '''		DTAS	(0x2)
    '''		DCAS	(0x1)
    '''	'RefData' must be a pinned IntPtr. Use GCHandle.Alloc to get a pinned pointer,
    '''	and pass handle.AddrOfPinnedObject() as 'RefData'.
    ''' </remarks>
    Public Shared Function ibnotify(ByVal ud As Integer, ByVal mask As Integer, ByVal Callback As GpibNotifyCallback_t, ByVal RefData As IntPtr) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibnotify32(ud, mask, Callback, RefData)
    End Function

    ''' <summary>
    ''' Reset an interface or device, or take it offline
    ''' </summary>
    ''' <param name="ud">Board (interface) or device descriptor</param>
    ''' <param name="v">0 to reset device/interface and take offline, 1 to just reset</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibonl(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibonl32(ud, v)
    End Function

    ''' <summary>
    ''' Change the primary GPIB address of a board or device
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to set</param>
    ''' <param name="v">New primary address (0 to 30)</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibpad(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibpad32(ud, v)
    End Function

    ''' <summary>
    ''' Pass Controller in Charge (CIC) status to another GPIB device
    ''' </summary>
    ''' <param name="ud">Device descriptor to make CIC</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibpct(ByVal ud As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibpct32(ud)
    End Function

    ''' <summary>
    ''' Not recommended for new code.
    ''' </summary>
    ''' <param name="ud"></param>
    ''' <param name="optionVal"></param>
    ''' <param name="v"></param>
    ''' <returns></returns>
    Public Shared Function ibpoke(ByVal ud As Integer, ByVal optionVal As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibpoke32(ud, optionVal, v)
    End Function

    ''' <summary>
    ''' Configure parallel polling
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to configure</param>
    ''' <param name="v">0 to disable, 96 to 126 to enable</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' 'v' values are from 96 to 126 (0x60 to 0x7E). Bits are encoded as follows:
    ''' 
    '''		0 1 1 E S D2 D1 D0
    '''		
    '''	where E=1 disables and E=0 enables parallel polling on this device,
    '''	S=1 asserts the data line when ist is 1, S=0 asserts the data line
    '''	when ist is 0, and D2 - D0 indicate which of the eight lines to assert. 	 
    ''' </remarks>
    Public Shared Function ibppc(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibppc32(ud, v)
    End Function

    ' Note for Function ibrd(ByVal ud as Integer, ByVal buf() as byte, ByVal cnt as Integer) As Integer
    '	This Method's "buf" parameter is passed directly to the API
    '	implementation. 
    ''' <summary>
    ''' Read data from the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to read from</param>
    ''' <param name="buf">Buffer to read data into</param>
    ''' <param name="cnt">Number of bytes to read</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibrd(ByVal ud As Integer, ByVal buf() As Byte, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrd32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Read data from the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to read from</param>
    ''' <param name="buf">Buffer to read data into</param>
    ''' <param name="cnt">Number of bytes to read</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibrd(ByVal ud As Integer, ByVal buf() As Short, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrd32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Read data from the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to read from</param>
    ''' <param name="buf">Buffer to read data into</param>
    ''' <param name="cnt">Number of bytes to read</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibrd(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrd32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Read data from the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to read from</param>
    ''' <param name="buf">Buffer to read data into</param>
    ''' <param name="cnt">Number of bytes to read</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibrd(ByVal ud As Integer, ByVal buf() As Single, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrd32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Read data from the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to read from</param>
    ''' <param name="buf">Buffer to read data into</param>
    ''' <param name="cnt">Number of bytes to read</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibrd(ByVal ud As Integer, ByVal buf() As Double, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrd32(ud, buf, cnt)
    End Function

    ' Note for Function ibrd(ByVal ud as Integer, ByVal buf as IntPtr, ByVal cnt as Integer) As Integer
    '	This Method's "buf" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "buf" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    ''' <summary>
    ''' Read data from the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to read from</param>
    ''' <param name="buf">Buffer to read data into (must be a pinned GCHandle pointer to a byte array)</param>
    ''' <param name="cnt">Number of bytes to read</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' 'buf' must be a pinned pointer to a byte array.
    ''' Use GCHandle.Alloc to get a GCHandle to the byte array,
    ''' and pass handle.AddrOfPinnedObject() as the 'buf' parameter.
    ''' </remarks>
    Public Shared Function ibrd(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrd32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Read data from the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to read from</param>
    ''' <param name="response">Buffer to read data into</param>
    ''' <param name="cnt">Number of characters to read</param>
    ''' <returns>Status outcome</returns>	
    Public Shared Function ibrd(ByVal ud As Integer, <Out()> ByRef response As String, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Dim buffer As New StringBuilder(cnt)
        Dim result As Integer = ibrd32(ud, buffer, cnt)
        response = buffer.ToString(0, ibcntl)
        Return result
    End Function

    Public Shared Function ibrda(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrda32(ud, buf, cnt)
    End Function

    ' Note: the buf parameter must point to a GCHandle to a pinned .NET
    ' * object (see GCHandle and .NET interoperation documentation for more
    ' * information) that will receive the data resulting from this command. 
    ''' <summary>
    ''' Read data from the GPIB asynchronously
    ''' </summary>
    ''' <param name="ud">Board or device descriptor to read from</param>
    ''' <param name="buf">Buffer to read data into (must be a pinned GCHandle pointer to a byte array)</param>
    ''' <param name="cnt">Number of bytes to read</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' 'buf' must be a pinned pointer to a byte array.
    ''' Use GCHandle.Alloc to get a GCHandle to the byte array,
    ''' and pass handle.AddrOfPinnedObject() as the 'buf' parameter.
    ''' </remarks>
    Public Shared Function ibrda(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrda32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Parallel poll all devices on an interface
    ''' </summary>
    ''' <param name="ud">Board (interface) or device descriptor</param>
    ''' <param name="ppr">Parallel poll assigned device bits</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' The meaning of the 'ppr' bits is determined by ibppc configuration calls.
    ''' </remarks>
    Public Shared Function ibrpp(ByVal ud As Integer, <Out()> ByRef ppr As Byte) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrpp32(ud, ppr)
    End Function

    ''' <summary>
    ''' Request or release System Controller status
    ''' </summary>
    ''' <param name="ud">Board (interface) descriptor</param>
    ''' <param name="v">Zero to release, nonzero to request</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' Agilent interfaces do not support changing this status.
    ''' If no error occurs, iberr contains the prior setting of this state.
    ''' </remarks>
    Public Shared Function ibrsc(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrsc32(ud, v)
    End Function

    ''' <summary>
    ''' Serial poll a device
    ''' </summary>
    ''' <param name="ud">Device descriptor to poll</param>
    ''' <param name="spr">Response byte. If 0x40 is set, the device is requesting service</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibrsp(ByVal ud As Integer, <Out()> ByRef spr As Byte) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrsp32(ud, spr)
    End Function

    ''' <summary>
    ''' Request service or change serial poll response byte
    ''' </summary>
    ''' <param name="ud">Board (interface) descriptor</param>
    ''' <param name="v">0x40 set requests service, other values change response byte</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibrsv(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibrsv32(ud, v)
    End Function

    ''' <summary>
    ''' Change the secondary GPIB address
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="v">Zero to disable SAD, 96 to 126 as new SAD value</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' Agilent interfaces don't support secondary addresses for themselves.
    ''' Agilent Connection Expert displays device secondary addresses as 0 to 30.
    ''' Add 96 to that value to get the equivalent Agilent 488 secondary address.
    ''' </remarks>
    Public Shared Function ibsad(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibsad32(ud, v)
    End Function

    ''' <summary>
    ''' Make the interface the Controller in Charge
    ''' </summary>
    ''' <param name="ud">Board (interface) descriptor</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibsic(ByVal ud As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibsic32(ud)
    End Function

    ''' <summary>
    ''' Assert or deassert Remote Enable (GPIB REN line)
    ''' </summary>
    ''' <param name="ud">Board (interface) descriptor</param>
    ''' <param name="v">Zero to deassert, nonzero to assert</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibsre(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibsre32(ud, v)
    End Function

    ''' <summary>
    ''' Stop an asynchronous I/O operation
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibstop(ByVal ud As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibstop32(ud)
    End Function

    ''' <summary>
    ''' Change the timeout period
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="v">Timeout period (Txxx constant; see IO Libraries Suite help on ibtmo)</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibtmo(ByVal ud As Integer, ByVal v As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibtmo32(ud, v)
    End Function

    ''' <summary>
    ''' Trigger a device.
    ''' </summary>
    ''' <param name="ud">Device descriptor</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibtrg(ByVal ud As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibtrg32(ud)
    End Function

    ''' <summary>
    ''' Wait for an asynchronous I/O operation event
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="mask">Event mask to wait for</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' Event constants for the event mask:
    '''		ERR		(0x8000)
    '''		TIMO	(0x4000)
    '''		END		(0x2000)
    '''		SRQI	(0x1000)
    '''		RQS		(0x800)
    '''		CMPL	(0x100)
    '''		LOK		(0x80)
    '''		REM		(0x40)
    '''		CIC		(0x20)
    '''		ATN		(0x10)
    '''		TACS	(0x8)
    '''		LACS	(0x4)
    '''		DTAS	(0x2)
    '''		DCAS	(0x1)
    '''	If TIMO is not set in the mask, ibwait will wait indefinitely for one of the masked events.
    '''	Set TIMO to ensure control returns no later that the ibtmo timeout period.
    '''	</remarks>
    Public Shared Function ibwait(ByVal ud As Integer, ByVal mask As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwait32(ud, mask)
    End Function

    ' Note for Function ibwrt(ByVal ud as Integer, ByVal buf() as byte, ByVal cnt as Integer) As Integer
    '	This Method's "buf" parameter is passed directly to the API
    '	implementation. 
    Public Shared Function ibwrt(ByVal ud As Integer, ByVal buf() As Byte, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrt32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Write data to the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="buf">Buffer to write</param>
    ''' <param name="cnt">Number of bytes to write</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibwrt(ByVal ud As Integer, ByVal buf() As Short, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrt32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Write data to the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="buf">Buffer to write</param>
    ''' <param name="cnt">Number of bytes to write</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibwrt(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrt32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Write data to the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="buf">Buffer to write</param>
    ''' <param name="cnt">Number of bytes to write</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibwrt(ByVal ud As Integer, ByVal buf() As Single, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrt32(ud, buf, cnt)
    End Function

    ''' <summary>
    ''' Write data to the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="buf">Buffer to write</param>
    ''' <param name="cnt">Number of bytes to write</param>
    ''' <returns>Status outcome</returns>
    Public Shared Function ibwrt(ByVal ud As Integer, ByVal buf() As Double, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrt32(ud, buf, cnt)
    End Function

    ' Note for Function ibwrt(ByVal ud as Integer, ByVal buf as IntPtr, ByVal cnt as Integer) As Integer
    '	This Method's "buf" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "buf" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    ''' <summary>
    ''' Write data to the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="buf">Buffer to write (must be a pinned GCHandle pointer to a byte array)</param>
    ''' <param name="cnt">Number of bytes to write</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' 'buf' must be a pinned pointer to a byte array.
    ''' Use GCHandle.Alloc to get a GCHandle to the byte array,
    ''' and pass handle.AddrOfPinnedObject() as the 'buf' parameter.
    ''' </remarks>
    Public Shared Function ibwrt(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrt32(ud, buf, cnt)
    End Function

    ' Note for Function ibwrt(ByVal ud as Integer, ByVal buf as string, ByVal cnt as Integer) as Integer
    '	This Method's "buf" parameter is treated as an ASCII string that 
    '	is converted by .NET into an array of bytes.  Consider using 
    '	a different version of this method for passing binary data. 
    ''' <summary>
    ''' Write data to the GPIB
    ''' </summary>
    ''' <param name="ud">Board or device descriptor</param>
    ''' <param name="buf">Buffer to write (characters only: see remarks)</param>
    ''' <param name="cnt">Number of bytes to write</param>
    ''' <returns>Status outcome</returns>
    ''' <remarks>
    ''' 'buf' is treated as an ASCII string, converted by .NET to an array of bytes.
    ''' Pass 'buf' as a byte array to send binary data.
    ''' </remarks>
    Public Shared Function ibwrt(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrt32(ud, buf, cnt)
    End Function

    Public Shared Function ibwrta(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrta32(ud, buf, cnt)
    End Function

    ' Note: the buf parameter must point to a GCHandle to a pinned .NET
    ' * object (ByVal GCHandle as see and .NET interoperation document for more
    ' * information) containing the data to send to the device 
    ''' <summary>
    ''' Write data asynchronously from a pinned GCHandle byte array to a device.
    ''' </summary>
    ''' <param name="ud">device or board descriptor</param>
    ''' <param name="buf">GCHandle pinned pointer to byte array</param>
    ''' <param name="cnt">bytes to write</param>
    ''' <returns>ibsta status</returns>
    ''' <remarks>
    ''' To use this method, obtain a pinned GCHandle using GCHandle.Alloc for a byte array, then pass handle.AddrOfPinnedObject() as 'buf'.
    ''' Use ibwait to wait for completion of this write before attempting further IO.
    ''' </remarks>
    Public Shared Function ibwrta(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrta32(ud, buf, cnt)
    End Function

    ' Note for Function ibwrta(ByVal ud as Integer, ByVal buf as string, ByVal cnt as Integer) As Integer
    '	This Method's "buf" parameter is treated as an ASCII string that 
    '	is converted by .NET into an array of bytes.  Consider using 
    '	a different version of this method for passing binary data. 
    Public Shared Function ibwrta(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibwrta32(ud, buf, cnt)
    End Function

    ' GPIB-ENET only functions to support locking across machines
    ' Deprecated - Use iblck

    ''' <summary>
    ''' Not recommended for new code. Use 'iblck'
    ''' </summary>
    ''' <param name="ud"></param>
    ''' <returns></returns>
    Public Shared Function iblock(ByVal ud As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return iblock32(ud)
    End Function

    ''' <summary>
    ''' Not recommended for new code. Use 'iblck'
    ''' </summary>
    ''' <param name="ud"></param>
    ''' <returns></returns>
    Public Shared Function ibunlock(ByVal ud As Integer) As Integer
        VerifyGpibGlobalsRegistered()
        Return ibunlock32(ud)
    End Function

    '************************************************************************
    '  Functions to access Thread-Specific copies of the GPIB global vars 


    ''' <summary>
    ''' Get the thread-specific ibsta value
    ''' </summary>
    ''' <returns>Thread-specific ibsta status</returns>
    ''' <remarks>
    ''' Use this instead of ibsta in multi-threaded applications
    ''' </remarks>
    Public Shared Function ThreadIbsta() As Integer
        VerifyGpibGlobalsRegistered()
        Return ThreadIbsta32()
    End Function

    ''' <summary>
    ''' Get the thread-specific iberr value
    ''' </summary>
    ''' <returns>Thread-specific iberr value</returns>
    ''' <remarks>
    ''' Use this instead of iberr in multi-threaded applications.
    ''' </remarks>
    Public Shared Function ThreadIberr() As Integer
        VerifyGpibGlobalsRegistered()
        Return ThreadIberr32()
    End Function

    ''' <summary>
    ''' Get the thread-specific ibcnt value
    ''' </summary>
    ''' <returns>Thread-specific ibcnt value</returns>
    ''' <remarks>
    ''' Use this instead of ibcnt in multi-threaded applications
    ''' </remarks>
    Public Shared Function ThreadIbcnt() As Integer
        VerifyGpibGlobalsRegistered()
        Return ThreadIbcnt32()
    End Function

    ''' <summary>
    ''' Get the thread-specific ibcntl value
    ''' </summary>
    ''' <returns>Thread-specific ibcntl value</returns>
    ''' <remarks>
    ''' Use this instead of ibcntl in multi-threaded applications
    ''' </remarks>
    Public Shared Function ThreadIbcntl() As Integer
        VerifyGpibGlobalsRegistered()
        Return ThreadIbcntl32()
    End Function


    '************************************************************************
    '  Agilent 488 Function Prototypes  


    ''' <summary>
    ''' Serial poll multiple devices
    ''' </summary>
    ''' <param name="boardID">Board number to poll</param>
    ''' <param name="addrlist">Array of device GPIB addresses to poll, ending with NOADDR</param>
    ''' <param name="results">Array of polling results, indexed like 'addrlist'</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub AllSpoll(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal results() As Short)
        VerifyGpibGlobalsRegistered()
        AllSpoll32(boardID, addrlist, results)
    End Sub

    ''' <summary>
    ''' Clear a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address of device</param>
    ''' <remarks>
    ''' GPIB address = sum of the primary (0-30) and secondary (0 or 96 - 126) addresses
    ''' </remarks>
    Public Shared Sub DevClear(ByVal boardID As Integer, ByVal addr As Short)
        VerifyGpibGlobalsRegistered()
        DevClear32(boardID, addr)
    End Sub

    ''' <summary>
    ''' Clear multiple devices
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses to clear, terminated with NOADDR</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub DevClearList(ByVal boardID As Integer, ByVal addrlist() As Short)
        VerifyGpibGlobalsRegistered()
        DevClearList32(boardID, addrlist)
    End Sub

    ''' <summary>
    ''' Place multiple devices in local mode
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses to make local, terminated with NOADDR</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' If 'addrlist' contains only NOADDR, all devices on the interface are local enabled.
    ''' </remarks>
    Public Shared Sub EnableLocal(ByVal boardID As Integer, ByVal addrlist() As Short)
        VerifyGpibGlobalsRegistered()
        EnableLocal32(boardID, addrlist)
    End Sub

    ''' <summary>
    ''' Place multiple devices in remote mode
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses to make remote, terminated by NOADDR</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' If 'addrlist' contains only NOADDR, all devices on the interface are remote enabled.
    ''' </remarks>
    Public Shared Sub EnableRemote(ByVal boardID As Integer, ByVal addrlist() As Short)
        VerifyGpibGlobalsRegistered()
        EnableRemote32(boardID, addrlist)
    End Sub

    ''' <summary>
    ''' Find listeners on the GPIB
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of primary addresses to search, terminated with NOADDR</param>
    ''' <param name="results">A list of the GPIB addresses of listening devices</param>
    ''' <param name="limit">Maximum number of listeners to return (should be the size of the results parameter)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub FindLstn(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal results() As Short, ByVal limit As Integer)
        VerifyGpibGlobalsRegistered()
        FindLstn32(boardID, addrlist, results, limit)
    End Sub

    ''' <summary>
    ''' Find a device requesting service
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses to search, terminated by NOADDR</param>
    ''' <param name="dev_stat">Serial poll response byte for the device requesting service</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' 'ibcntl' contains the index of the device requesting service, or if none, the index of the terminating NOADDR.
    ''' If no device requests service, 'iberr' contains ETAB.
    ''' </remarks>
    Public Shared Sub FindRQS(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal dev_stat() As Short)
        VerifyGpibGlobalsRegistered()
        FindRQS32(boardID, addrlist, dev_stat)
    End Sub

    ''' <summary>
    ''' Perform a parallel poll
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="result">8-bit poll result, with 1-bit status per device </param>
    Public Shared Sub PPoll(ByVal boardID As Integer, <Out()> ByRef result As Short)
        VerifyGpibGlobalsRegistered()
        PPoll32(boardID, result)
    End Sub

    ''' <summary>
    ''' Configure a device for parallel poll
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address of device to configure</param>
    ''' <param name="dataLine">Line number (1 - 8) assigned to this device</param>
    ''' <param name="lineSense">Response sense: 1 to assert, 0 to deassert</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub PPollConfig(ByVal boardID As Integer, ByVal addr As Short, ByVal dataLine As Integer, ByVal lineSense As Integer)
        VerifyGpibGlobalsRegistered()
        PPollConfig32(boardID, addr, dataLine, lineSense)
    End Sub

    ''' <summary>
    ''' Unconfigure devices so they don't respond to parallel polls
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses to unconfigure, terminated by NOADDR</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub PPollUnconfig(ByVal boardID As Integer, ByVal addrlist() As Short)
        VerifyGpibGlobalsRegistered()
        PPollUnconfig32(boardID, addrlist)
    End Sub

    ''' <summary>
    ''' Pass Controller in Charge status to a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub PassControl(ByVal boardID As Integer, ByVal addr As Short)
        VerifyGpibGlobalsRegistered()
        PassControl32(boardID, addr)
    End Sub

    ' Note for Function void RcvRespMsg(ByVal boardID as Integer, ByVal buffer() as byte, ByVal cnt as Integer, ByVal Termination as Integer)
    '	This method's "buffer" parameter is passed directly to the API
    '	implementation. 
    ''' <summary>
    ''' Read data from an addressed device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Buffer to hold data</param>
    ''' <param name="cnt">Maximum bytes to read</param>
    ''' <param name="Termination">Either a termination character, or STOPend to terminate on EOI</param>
    ''' <remarks>
    ''' Use 'ReceiveSetup' to address a device before calling this method.
    ''' </remarks>
    Public Shared Sub RcvRespMsg(ByVal boardID As Integer, ByVal buffer() As Byte, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        RcvRespMsg32(boardID, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Read data from an addressed device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Buffer to hold data</param>
    ''' <param name="cnt">Maximum bytes to read</param>
    ''' <param name="Termination">Either a termination character, or STOPend to terminate on EOI</param>
    ''' <remarks>
    ''' Use 'ReceiveSetup' to address a device before calling this method.
    ''' </remarks>
    Public Shared Sub RcvRespMsg(ByVal boardID As Integer, ByVal buffer() As Short, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        RcvRespMsg32(boardID, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Read data from an addressed device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Buffer to hold data</param>
    ''' <param name="cnt">Maximum bytes to read</param>
    ''' <param name="Termination">Either a termination character, or STOPend to terminate on EOI</param>
    ''' <remarks>
    ''' Use 'ReceiveSetup' to address a device before calling this method.
    ''' </remarks>
    Public Shared Sub RcvRespMsg(ByVal boardID As Integer, ByVal buffer() As Integer, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        RcvRespMsg32(boardID, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Read data from an addressed device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Buffer to hold data</param>
    ''' <param name="cnt">Maximum bytes to read</param>
    ''' <param name="Termination">Either a termination character, or STOPend to terminate on EOI</param>
    ''' <remarks>
    ''' Use 'ReceiveSetup' to address a device before calling this method.
    ''' </remarks>
    Public Shared Sub RcvRespMsg(ByVal boardID As Integer, ByVal buffer() As Single, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        RcvRespMsg32(boardID, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Read data from an addressed device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Buffer to hold data</param>
    ''' <param name="cnt">Maximum bytes to read</param>
    ''' <param name="Termination">Either a termination character, or STOPend to terminate on EOI</param>
    ''' <remarks>
    ''' Use 'ReceiveSetup' to address a device before calling this method.
    ''' </remarks>
    Public Shared Sub RcvRespMsg(ByVal boardID As Integer, ByVal buffer() As Double, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        RcvRespMsg32(boardID, buffer, cnt, Termination)
    End Sub

    ' Note for Function void RcvRespMsg(ByVal boardID as Integer, ByVal buffer as IntPtr, ByVal cnt as Integer, ByVal Termination as Integer)
    '	This Method's "buffer" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "buffer" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    ''' <summary>
    ''' Read data from an addressed device.
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">GCHandle pinned pointer to a byte array to hold data</param>
    ''' <param name="cnt">Maximum bytes to read</param>
    ''' <param name="Termination">Either a termination character, or STOPend to terminate on EOI</param>
    ''' <remarks>
    ''' Use 'ReceiveSetup' to address a device before calling this method.
    ''' Use GCHandle.Alloc to get a pinned pointer to the byte array, and
    ''' pass handle.AddrOfPinnedObject() as 'buffer'
    ''' </remarks>
    Public Shared Sub RcvRespMsg(ByVal boardID As Integer, ByVal buffer As IntPtr, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        RcvRespMsg32(boardID, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Read data from an addressed device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="response">String to hold data</param>
    ''' <param name="cnt">Maximum characters to read</param>
    ''' <param name="Termination">Either a termination character, or STOPend to terminate on EOI</param>
    ''' <remarks>
    ''' Use 'ReceiveSetup' to address a device before calling this method.
    ''' </remarks>
    Public Shared Sub RcvRespMsg(ByVal boardID As Integer, <Out()> ByRef response As String, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        Dim buffer As New StringBuilder(cnt)
        RcvRespMsg32(boardID, buffer, cnt, Termination)
        response = buffer.ToString(0, ibcntl)
    End Sub

    ''' <summary>
    ''' Serial poll a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address of device</param>
    ''' <param name="result">Serial poll response byte from the device</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub ReadStatusByte(ByVal boardID As Integer, ByVal addr As Short, <Out()> ByRef result As Short)
        VerifyGpibGlobalsRegistered()
        ReadStatusByte32(boardID, addr, result)
    End Sub

    ' Note for Function void Receive(ByVal boardID as Integer, ByVal addr as short, ByVal buffer() as byte, ByVal cnt as Integer, ByVal Termination as Integer)
    '	This Method's "buffer" parameter is passed directly to the API
    '	implementation. 
    ''' <summary>
    ''' Receive data from a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="buffer">Buffer to receive data</param>
    ''' <param name="cnt">Maximum number of bytes to receive</param>
    ''' <param name="Termination">Termination character to stop after, or STOPend to stop on EOI</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Receive(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Byte, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        Receive32(boardID, addr, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Receive data from a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="buffer">Buffer to receive data</param>
    ''' <param name="cnt">Maximum number of bytes to receive</param>
    ''' <param name="Termination">Termination character to stop after, or STOPend to stop on EOI</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Receive(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Short, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        Receive32(boardID, addr, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Receive data from a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="buffer">Buffer to receive data</param>
    ''' <param name="cnt">Maximum number of bytes to receive</param>
    ''' <param name="Termination">Termination character to stop after, or STOPend to stop on EOI</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Receive(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Integer, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        Receive32(boardID, addr, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Receive data from a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="buffer">Buffer to receive data</param>
    ''' <param name="cnt">Maximum number of bytes to receive</param>
    ''' <param name="Termination">Termination character to stop after, or STOPend to stop on EOI</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Receive(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Single, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        Receive32(boardID, addr, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Receive data from a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="buffer">Buffer to receive data</param>
    ''' <param name="cnt">Maximum number of bytes to receive</param>
    ''' <param name="Termination">Termination character to stop after, or STOPend to stop on EOI</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Receive(ByVal boardID As Integer, ByVal addr As Short, ByVal buffer() As Double, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        Receive32(boardID, addr, buffer, cnt, Termination)
    End Sub

    ''' <summary>
    ''' Receive data from a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="response">String to receive data</param>
    ''' <param name="cnt">Maximum number of characters to receive</param>
    ''' <param name="Termination">Termination character to stop after, or STOPend to stop on EOI</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Receive(ByVal boardID As Integer, ByVal addr As Short, <Out()> ByRef response As String, ByVal cnt As Integer, ByVal Termination As Integer)
        VerifyGpibGlobalsRegistered()
        Dim buffer As New StringBuilder(cnt)
        Receive32(boardID, addr, buffer, cnt, Termination)
        response = buffer.ToString(0, ibcntl)
    End Sub

    ''' <summary>
    ''' Address a device to receive data
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' Use this method to address a device before using RcvRespMsg to receive the data
    ''' </remarks>
    Public Shared Sub ReceiveSetup(ByVal boardID As Integer, ByVal addr As Short)
        VerifyGpibGlobalsRegistered()
        ReceiveSetup32(boardID, addr)
    End Sub

    ''' <summary>
    ''' Reset the GPIB system
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses to reset, terminated by NOADDR</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub ResetSys(ByVal boardID As Integer, ByVal addrlist() As Short)
        VerifyGpibGlobalsRegistered()
        ResetSys32(boardID, addrlist)
    End Sub

    ' Note for Function void Send(ByVal boardID as Integer, ByVal addr as short, ByVal databuf() as byte, ByVal datacnt as Integer, ByVal eotMode as Integer)
    '	This Method's "databuf" parameter is passed directly to the API
    '	implementation. 
    ''' <summary>
    ''' Send data to a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Send(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Byte, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        Send32(boardID, addr, databuf, datacnt, eotMode)
    End Sub

    ''' <summary>
    ''' Send data to a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Send(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Short, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        Send32(boardID, addr, databuf, datacnt, eotMode)
    End Sub

    ''' <summary>
    ''' Send data to a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Send(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Integer, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        Send32(boardID, addr, databuf, datacnt, eotMode)
    End Sub

    ''' <summary>
    ''' Send data to a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Send(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Single, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        Send32(boardID, addr, databuf, datacnt, eotMode)
    End Sub

    ''' <summary>
    ''' Send data to a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Send(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf() As Double, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        Send32(boardID, addr, databuf, datacnt, eotMode)
    End Sub

    ' Note for Function void Send(ByVal boardID as Integer, ByVal addr as short, ByVal databuf as IntPtr, ByVal datacnt as Integer, ByVal eotMode as Integer)
    '	This Method's "databuf" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "databuf" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    ''' <summary>
    ''' Send data to a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="databuf">GCHandle pinned pointer to byte array with data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' Use GCHandle.Alloc to get a pinned pointer to the byte array, and
    ''' pass handle.AddrOfPinnedObject() as 'databuf'
    ''' </remarks>
    Public Shared Sub Send(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf As IntPtr, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        Send32(boardID, addr, databuf, datacnt, eotMode)
    End Sub

    ' Note for Function void Send(ByVal boardID as Integer, ByVal addr as short, ByVal databuf as string, ByVal datacnt as Integer, ByVal eotMode as Integer)
    '	This Method's "databuf" parameter is treated as an ASCII string that 
    '	is converted by .NET into an array of bytes.  Consider using 
    '	a different version of this method for passing binary data. 
    ''' <summary>
    ''' Send data to a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' Send only character data with this method.  Send binary data using a byte array.
    ''' </remarks>
    Public Shared Sub Send(ByVal boardID As Integer, ByVal addr As Short, ByVal databuf As String, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        Send32(boardID, addr, databuf, datacnt, eotMode)
    End Sub

    ' Note for Function void SendCmds(ByVal boardID as Integer, ByVal buffer() as byte, ByVal cnt as Integer)
    '	This Method's "buffer" parameter is passed directly to the API
    '	implementation. 
    ''' <summary>
    ''' Send GPIB commands (interface messages) on the interface
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Commands to send</param>
    ''' <param name="cnt">Count of bytes to send</param>
    ''' <remarks>
    ''' Interface messages are one byte or character commands.  
    ''' See online help for possible command constants and character command meanings.
    ''' </remarks>
    Public Shared Sub SendCmds(ByVal boardID As Integer, ByVal buffer() As Byte, ByVal cnt As Integer)
        VerifyGpibGlobalsRegistered()
        SendCmds32(boardID, buffer, cnt)
    End Sub

    ''' <summary>
    ''' Not recommended for new code.
    ''' </summary>
    ''' <param name="boardID"></param>
    ''' <param name="buffer"></param>
    ''' <param name="cnt"></param>
    Public Shared Sub SendCmds(ByVal boardID As Integer, ByVal buffer() As Short, ByVal cnt As Integer)
        VerifyGpibGlobalsRegistered()
        SendCmds32(boardID, buffer, cnt)
    End Sub

    ''' <summary>
    ''' Not recommended for new code.
    ''' </summary>
    ''' <param name="boardID"></param>
    ''' <param name="buffer"></param>
    ''' <param name="cnt"></param>
    Public Shared Sub SendCmds(ByVal boardID As Integer, ByVal buffer() As Integer, ByVal cnt As Integer)
        VerifyGpibGlobalsRegistered()
        SendCmds32(boardID, buffer, cnt)
    End Sub

    ''' <summary>
    ''' Not recommended for new code.
    ''' </summary>
    ''' <param name="boardID"></param>
    ''' <param name="buffer"></param>
    ''' <param name="cnt"></param>
    Public Shared Sub SendCmds(ByVal boardID As Integer, ByVal buffer() As Single, ByVal cnt As Integer)
        VerifyGpibGlobalsRegistered()
        SendCmds32(boardID, buffer, cnt)
    End Sub

    ''' <summary>
    ''' Not recommended for new code.
    ''' </summary>
    ''' <param name="boardID"></param>
    ''' <param name="buffer"></param>
    ''' <param name="cnt"></param>
    Public Shared Sub SendCmds(ByVal boardID As Integer, ByVal buffer() As Double, ByVal cnt As Integer)
        VerifyGpibGlobalsRegistered()
        SendCmds32(boardID, buffer, cnt)
    End Sub

    ' Note for Function void SendCmds(ByVal boardID as Integer, ByVal buffer as IntPtr, ByVal cnt as Integer)
    '	This Method's "buffer" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "buffer" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    ''' <summary>
    ''' Not recommended for new code.
    ''' </summary>
    ''' <param name="boardID"></param>
    ''' <param name="buffer"></param>
    ''' <param name="cnt"></param>
    Public Shared Sub SendCmds(ByVal boardID As Integer, ByVal buffer As IntPtr, ByVal cnt As Integer)
        VerifyGpibGlobalsRegistered()
        SendCmds32(boardID, buffer, cnt)
    End Sub

    ' Note for Function void SendCmds(ByVal boardID as Integer, ByVal buffer as string, ByVal cnt as Integer)
    '	This Method's "buffer" parameter is treated as an ASCII string that 
    '	is converted by .NET into an array of bytes.  Consider using 
    '	a different version of this method for passing binary data. 
    ''' <summary>
    ''' Send GPIB commands (interface messages) on the interface
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Commands to send</param>
    ''' <param name="cnt">Count of characters to send</param>
    ''' <remarks>
    ''' Interface messages are one-byte or character commands.  
    ''' See online help for possible command constants and ASCII character command meanings.
    ''' </remarks>
    Public Shared Sub SendCmds(ByVal boardID As Integer, ByVal buffer As String, ByVal cnt As Integer)
        VerifyGpibGlobalsRegistered()
        SendCmds32(boardID, buffer, cnt)
    End Sub

    ' Note for Function void SendDataBytes(ByVal boardID as Integer, ByVal buffer() as byte, ByVal cnt as Integer, ByVal eot_mode as Integer)
    '	This Method's "buffer" parameter is passed directly to the API
    '	implementation. 
    ''' <summary>
    ''' Send data over GPIB
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Data to send</param>
    ''' <param name="cnt">Number of bytes to send</param>
    ''' <param name="eot_mode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' Use 'SendSetup' before this method to address the data recipient device.
    ''' </remarks>
    Public Shared Sub SendDataBytes(ByVal boardID As Integer, ByVal buffer() As Byte, ByVal cnt As Integer, ByVal eot_mode As Integer)
        VerifyGpibGlobalsRegistered()
        SendDataBytes32(boardID, buffer, cnt, eot_mode)
    End Sub

    ''' <summary>
    ''' Send data over GPIB
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Data to send</param>
    ''' <param name="cnt">Number of bytes to send</param>
    ''' <param name="eot_mode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' Use 'SendSetup' before this method to address the data recipient device.
    ''' </remarks>
    Public Shared Sub SendDataBytes(ByVal boardID As Integer, ByVal buffer() As Short, ByVal cnt As Integer, ByVal eot_mode As Integer)
        VerifyGpibGlobalsRegistered()
        SendDataBytes32(boardID, buffer, cnt, eot_mode)
    End Sub

    ''' <summary>
    ''' Send data over GPIB
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Data to send</param>
    ''' <param name="cnt">Number of bytes to send</param>
    ''' <param name="eot_mode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' Use 'SendSetup' before this method to address the data recipient device.
    ''' </remarks>
    Public Shared Sub SendDataBytes(ByVal boardID As Integer, ByVal buffer() As Integer, ByVal cnt As Integer, ByVal eot_mode As Integer)
        VerifyGpibGlobalsRegistered()
        SendDataBytes32(boardID, buffer, cnt, eot_mode)
    End Sub

    ''' <summary>
    ''' Send data over GPIB
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Data to send</param>
    ''' <param name="cnt">Number of bytes to send</param>
    ''' <param name="eot_mode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' Use 'SendSetup' before this method to address the data recipient device.
    ''' </remarks>
    Public Shared Sub SendDataBytes(ByVal boardID As Integer, ByVal buffer() As Single, ByVal cnt As Integer, ByVal eot_mode As Integer)
        VerifyGpibGlobalsRegistered()
        SendDataBytes32(boardID, buffer, cnt, eot_mode)
    End Sub

    ''' <summary>
    ''' Send data over GPIB
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Data to send</param>
    ''' <param name="cnt">Number of bytes to send</param>
    ''' <param name="eot_mode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' Use 'SendSetup' before this method to address the data recipient device.
    ''' </remarks>
    Public Shared Sub SendDataBytes(ByVal boardID As Integer, ByVal buffer() As Double, ByVal cnt As Integer, ByVal eot_mode As Integer)
        VerifyGpibGlobalsRegistered()
        SendDataBytes32(boardID, buffer, cnt, eot_mode)
    End Sub

    ' Note for Function void SendDataBytes(ByVal boardID as Integer, ByVal buffer as IntPtr, ByVal cnt as Integer, ByVal eot_mode as Integer)
    '	This Method's "buffer" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "buffer" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    ''' <summary>
    ''' Send data over GPIB
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">GCHandle pinned pointer to a byte array of data to send</param>
    ''' <param name="cnt">Number of bytes to send</param>
    ''' <param name="eot_mode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' Use 'SendSetup' before this method to address the data recipient device.
    ''' Use GCHandle.Alloc to get a pinned pointer to a byte array, and
    ''' pass handle.AddrOfPinnedObject() as 'buffer'.
    ''' </remarks>
    Public Shared Sub SendDataBytes(ByVal boardID As Integer, ByVal buffer As IntPtr, ByVal cnt As Integer, ByVal eot_mode As Integer)
        VerifyGpibGlobalsRegistered()
        SendDataBytes32(boardID, buffer, cnt, eot_mode)
    End Sub

    ' Note for Function void SendDataBytes(ByVal boardID as Integer, ByVal buffer as string, ByVal cnt as Integer, ByVal eot_mode as Integer)
    '	This Method's "buffer" parameter is treated as an ASCII string that 
    '	is converted by .NET into an array of bytes.  Consider using 
    '	a different version of this method for passing binary data. 
    ''' <summary>
    ''' Send character data over GPIB
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="buffer">Data to send</param>
    ''' <param name="cnt">Number of characters to send</param>
    ''' <param name="eot_mode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' Use 'SendSetup' before this method to address the data recipient device.
    ''' Use this form for character data, use a byte array for binary data.
    ''' </remarks>
    Public Shared Sub SendDataBytes(ByVal boardID As Integer, ByVal buffer As String, ByVal cnt As Integer, ByVal eot_mode As Integer)
        VerifyGpibGlobalsRegistered()
        SendDataBytes32(boardID, buffer, cnt, eot_mode)
    End Sub

    ''' <summary>
    ''' Clear the interface
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    Public Shared Sub SendIFC(ByVal boardID As Integer)
        VerifyGpibGlobalsRegistered()
        SendIFC32(boardID)
    End Sub

    ''' <summary>
    ''' Send Local Lockout to all devices on the interface
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    Public Shared Sub SendLLO(ByVal boardID As Integer)
        VerifyGpibGlobalsRegistered()
        SendLLO32(boardID)
    End Sub

    ' Note for Function void SendList(ByVal boardID as Integer, ByVal addrlist() as short, ByVal databuf() as byte, ByVal datacnt as Integer, ByVal eotMode as Integer)
    '	This Method's "databuf" parameter is passed directly to the API
    '	implementation. 
    ''' <summary>
    ''' Send data to a list of devices
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub SendList(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Byte, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        SendList32(boardID, addrlist, databuf, datacnt, eotMode)
    End Sub

    ''' <summary>
    ''' Send data to a list of devices
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub SendList(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Short, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        SendList32(boardID, addrlist, databuf, datacnt, eotMode)
    End Sub

    ''' <summary>
    ''' Send data to a list of devices
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub SendList(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Integer, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        SendList32(boardID, addrlist, databuf, datacnt, eotMode)
    End Sub

    ''' <summary>
    ''' Send data to a list of devices
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub SendList(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Single, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        SendList32(boardID, addrlist, databuf, datacnt, eotMode)
    End Sub

    ''' <summary>
    ''' Send data to a list of devices
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub SendList(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf() As Double, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        SendList32(boardID, addrlist, databuf, datacnt, eotMode)
    End Sub

    ' Note for Function void SendList(ByVal boardID as Integer, ByVal addrlist() as short, ByVal databuf as IntPtr, ByVal datacnt as Integer, ByVal eotMode as Integer)
    '	This Method's "databuf" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "databuf" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    ''' <summary>
    ''' Send data to a list of devices
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <param name="databuf">GCHandle pinned pointer to byte array with data to send</param>
    ''' <param name="datacnt">Number of bytes to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' Use GCHandle.Alloc to get a pinned pointer to a byte array, and
    ''' pass handle.AddrOfPinnedObject() as 'databuf'.
    ''' </remarks>
    Public Shared Sub SendList(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf As IntPtr, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        SendList32(boardID, addrlist, databuf, datacnt, eotMode)
    End Sub

    ' Note for Function void SendList(ByVal boardID as Integer, ByVal addrlist() as short, ByVal databuf as string, ByVal datacnt as Integer, ByVal eotMode as Integer)
    '	This Method's "databuf" parameter is treated as an ASCII string that 
    '	is converted by .NET into an array of bytes.  Consider using 
    '	a different version of this method for passing binary data. 
    ''' <summary>
    ''' Send data to a list of devices
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <param name="databuf">Data to send</param>
    ''' <param name="datacnt">Number of characters to send</param>
    ''' <param name="eotMode">End indicator: DABend (EOI), NLend (newline+EOI), or NULLend (none)</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' Use this for character data.  Use a byte array for binary data.
    ''' </remarks>
    Public Shared Sub SendList(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal databuf As String, ByVal datacnt As Integer, ByVal eotMode As Integer)
        VerifyGpibGlobalsRegistered()
        SendList32(boardID, addrlist, databuf, datacnt, eotMode)
    End Sub

    ''' <summary>
    ''' Address a device for send operations
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">GPIB address</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' Use this method a device before using 'SendCmds' or 'SendDataBytes'
    ''' </remarks>
    Public Shared Sub SendSetup(ByVal boardID As Integer, ByVal addrlist() As Short)
        VerifyGpibGlobalsRegistered()
        SendSetup32(boardID, addrlist)
    End Sub

    ''' <summary>
    ''' Put devices in remote state with local lockout and address them as Listeners
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub SetRWLS(ByVal boardID As Integer, ByVal addrlist() As Short)
        VerifyGpibGlobalsRegistered()
        SetRWLS32(boardID, addrlist)
    End Sub

    ''' <summary>
    ''' Return the state of the SRQ line
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="result">1 if SRQ asserted, 0 if deasserted</param>
    Public Shared Sub TestSRQ(ByVal boardID As Integer, <Out()> ByRef result As Short)
        VerifyGpibGlobalsRegistered()
        TestSRQ32(boardID, result)
    End Sub

    ''' <summary>
    ''' Send self-test message to listed devices and return self-test results
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <param name="results">Array of self-test results: usually 0 means no error</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub TestSys(ByVal boardID As Integer, ByVal addrlist() As Short, ByVal results() As Short)
        VerifyGpibGlobalsRegistered()
        TestSys32(boardID, addrlist, results)
    End Sub

    ''' <summary>
    ''' Trigger a device
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addr">GPIB address</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub Trigger(ByVal boardID As Integer, ByVal addr As Short)
        VerifyGpibGlobalsRegistered()
        Trigger32(boardID, addr)
    End Sub

    ''' <summary>
    ''' Trigger a list of devices
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="addrlist">Array of GPIB addresses, terminated by NOADDR</param>
    ''' <remarks>
    ''' GPIB address has a low byte primary address (0-30) and a high byte secondary address (0 or 96 - 126)
    ''' </remarks>
    Public Shared Sub TriggerList(ByVal boardID As Integer, ByVal addrlist() As Short)
        VerifyGpibGlobalsRegistered()
        TriggerList32(boardID, addrlist)
    End Sub

    ''' <summary>
    ''' Wait for SRQ to be asserted
    ''' </summary>
    ''' <param name="boardID">Interface board number</param>
    ''' <param name="result">1 if SRQ asserted before this method returns, 0 otherwise</param>
    Public Shared Sub WaitSRQ(ByVal boardID As Integer, <Out()> ByRef result As Short)
        VerifyGpibGlobalsRegistered()
        WaitSRQ32(boardID, result)
    End Sub
#End Region

#Region "Deprecated Function versions for VB6 upgraders"
    ' The functions in this section below are for compatibility with older Visual Basic 6 programs
    ' that were upgraded to .NET.  These functions provide duplicate functionality to 
    ' already-defined functions and are named differently as a result of type and method 
    ' overloading limitations of Visual Basic 6 that went away with VB .NET

    Public Shared Function ibrdi(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        Return ibrd(ud, buf, cnt)
    End Function

    Public Shared Function ibrdia(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        Return ibrda(ud, buf, cnt)
    End Function

    Public Shared Function ibwrti(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        Return ibwrt(ud, buf, cnt)
    End Function

    Public Shared Function ibwrtia(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        Return ibwrta(ud, buf, cnt)
    End Function

    Public Shared Function ilask(ByVal ud As Integer, ByVal opt As Integer, <Out()> ByRef rval As Integer) As Integer
        Return ibask(ud, opt, rval)
    End Function

    Public Shared Function ilbna(ByVal ud As Integer, ByVal udname As String) As Integer
        Return ibbna(ud, udname)
    End Function

    Public Shared Function ilcac(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibcac(ud, v)
    End Function

    Public Shared Function ilclr(ByVal ud As Integer) As Integer
        Return ibclr(ud)
    End Function

    Public Shared Function ilcmd(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
        Return ibcmd(ud, buf, cnt)
    End Function

    Public Shared Function ilcmda(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
        Return ibcmda(ud, buf, cnt)
    End Function

    Public Shared Function ilconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal v As Integer) As Integer
        Return ibconfig(bdid, opt, v)
    End Function

    Public Shared Function ildev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
        Return ibdev(bdid, pad, sad, tmo, eot, eos)
    End Function

    Public Shared Function ildma(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibdma(ud, v)
    End Function

    Public Shared Function ileos(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibeos(ud, v)
    End Function

    Public Shared Function ileot(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibeot(ud, v)
    End Function

    Public Shared Function ilfind(ByVal udname As String) As Integer
        Return ibfind(udname)
    End Function

    Public Shared Function ilgts(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibgts(ud, v)
    End Function

    Public Shared Function ilist(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibist(ud, v)
    End Function

    Public Shared Function illck(ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Integer) As Integer
        Return iblck(ud, v, LockWaitTime, IntPtr.Zero)
    End Function

    Public Shared Function illines(ByVal ud As Integer, <Out()> ByRef lines As Short) As Integer
        Return iblines(ud, lines)
    End Function

    Public Shared Function illn(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, <Out()> ByRef listen As Short) As Integer
        Return ibln(ud, pad, sad, listen)
    End Function

    Public Shared Function illoc(ByVal ud As Integer) As Integer
        Return ibloc(ud)
    End Function

    Public Shared Function ilonl(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibonl(ud, v)
    End Function

    Public Shared Function ilpad(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibpad(ud, v)
    End Function

    Public Shared Function ilpct(ByVal ud As Integer) As Integer
        Return ibpct(ud)
    End Function

    Public Shared Function ilppc(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibppc(ud, v)
    End Function

    ' Note for Function ibrd(ByVal ud as Integer, ByVal buf() as byte, ByVal cnt as Integer) as Integer
    '	This Method's "buf" parameter is passed directly to the API
    '	implementation. 
    Public Shared Function ilrd(ByVal ud As Integer, ByVal buf() As Byte, ByVal cnt As Integer) As Integer
        Return ibrd(ud, buf, cnt)
    End Function

    Public Shared Function ilrd(ByVal ud As Integer, ByVal buf() As Short, ByVal cnt As Integer) As Integer
        Return ibrd(ud, buf, cnt)
    End Function

    Public Shared Function ilrd(ByVal ud As Integer, ByVal buf() As Integer, ByVal cnt As Integer) As Integer
        Return ibrd(ud, buf, cnt)
    End Function

    Public Shared Function ilrd(ByVal ud As Integer, ByVal buf() As Single, ByVal cnt As Integer) As Integer
        Return ibrd(ud, buf, cnt)
    End Function

    Public Shared Function ilrd(ByVal ud As Integer, ByVal buf() As Double, ByVal cnt As Integer) As Integer
        Return ibrd(ud, buf, cnt)
    End Function

    ' Note for Function ilrd(ByVal ud as Integer, ByVal buf as IntPtr, ByVal cnt as Integer) as Integer
    '	This Method's "buf" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "buf" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    Public Shared Function ilrd(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
        Return ibrd(ud, buf, cnt)
    End Function

    Public Shared Function ilrd(ByVal ud As Integer, <Out()> ByRef response As String, ByVal cnt As Integer) As Integer
        Return ibrd(ud, response, cnt)
    End Function

    Public Shared Function ilrdi(ByVal ud As Integer, ByVal ibuf() As Integer, ByVal cnt As Integer) As Integer
        Return ibrdi(ud, ibuf, cnt)
    End Function


    ' Note for Function ilrda(ByVal ud as Integer, ByVal buf as IntPtr, ByVal cnt as Integer) as Integer
    '	This Method's "buf" parameter is passed directly to .NET's unmanaged 
    '	interoperability layer.  You are responsible for initializing "buf" 
    '	as a valid pinned pointer to managed memory when you use this 
    '	overload of the method. 
    Public Shared Function ilrda(ByVal ud As Integer, ByVal buf As IntPtr, ByVal cnt As Integer) As Integer
        Return ibrda(ud, buf, cnt)
    End Function

    Public Shared Function ilrdf(ByVal ud As Integer, ByVal filename As String) As Integer
        Return ibrdf(ud, filename)
    End Function

    Public Shared Function ilrdia(ByVal ud As Integer, ByVal ibuf() As Integer, ByVal cnt As Integer) As Integer
        Return ibrdia(ud, ibuf, cnt)
    End Function

    Public Shared Function ilrpp(ByVal ud As Integer, <Out()> ByRef ppr As Byte) As Integer
        Return ibrpp(ud, ppr)
    End Function

    Public Shared Function ilrsc(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibrsc(ud, v)
    End Function

    Public Shared Function ilrsp(ByVal ud As Integer, <Out()> ByRef spr As Byte) As Integer
        Return ibrsp(ud, spr)
    End Function

    Public Shared Function ilrsv(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibrsv(ud, v)
    End Function

    Public Shared Function ilsad(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibsad(ud, v)
    End Function

    Public Shared Function ilsic(ByVal ud As Integer) As Integer
        Return ibsic(ud)
    End Function

    Public Shared Function ilsre(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibsre(ud, v)
    End Function

    Public Shared Function ilstop(ByVal ud As Integer) As Integer
        Return ibstop(ud)
    End Function

    Public Shared Function iltmo(ByVal ud As Integer, ByVal v As Integer) As Integer
        Return ibtmo(ud, v)
    End Function

    Public Shared Function iltrg(ByVal ud As Integer) As Integer
        Return ibtrg(ud)
    End Function

    Public Shared Function ilwait(ByVal ud As Integer, ByVal mask As Integer) As Integer
        Return ibwait(ud, mask)
    End Function

    Public Shared Function ilwrt(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
        Return ibwrt(ud, buf, cnt)
    End Function

    Public Shared Function ilwrta(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Integer) As Integer
        Return ibwrta(ud, buf, cnt)
    End Function

    Public Shared Function ilwrtf(ByVal ud As Integer, ByVal filename As String) As Integer
        Return ibwrtf(ud, filename)
    End Function

    Public Shared Function ilwrti(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Integer) As Integer
        Return ibwrti(ud, ibuf, cnt)
    End Function

    Public Shared Function ilwrtia(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Integer) As Integer
        Return ibwrtia(ud, ibuf, cnt)
    End Function

#End Region

End Class

'end namespace
