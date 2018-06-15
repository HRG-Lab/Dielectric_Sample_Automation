Attribute VB_Name = "VBIB32"
' Copyright (C) 2005 Keysight Technologies, Inc.
'
' This file defines constants, record types, and entry points for the Keysight 488 API.
' You need to add this file to each Microsoft Visual Basic project that uses the
' Agilent 488 API
'
Option Explicit

Declare Function ibask32 Lib "Gpib-32.dll" Alias "ibask" (ByVal ud As Long, ByVal opt As Long, value As Long) As Long
Declare Function ibbna32 Lib "Gpib-32.dll" Alias "ibbnaA" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibcac32 Lib "Gpib-32.dll" Alias "ibcac" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibclr32 Lib "Gpib-32.dll" Alias "ibclr" (ByVal ud As Long) As Long
Declare Function ibcmd32 Lib "Gpib-32.dll" Alias "ibcmd" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibcmda32 Lib "Gpib-32.dll" Alias "ibcmda" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibconfig32 Lib "Gpib-32.dll" Alias "ibconfig" (ByVal ud As Long, ByVal opt As Long, ByVal v As Long) As Long
Declare Function ibdev32 Lib "Gpib-32.dll" Alias "ibdev" (ByVal bdid As Long, ByVal pad As Long, ByVal sad As Long, ByVal tmo As Long, ByVal eot As Long, ByVal eos As Long) As Long
Declare Function ibdma32 Lib "Gpib-32.dll" Alias "ibdma" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibeos32 Lib "Gpib-32.dll" Alias "ibeos" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibeot32 Lib "Gpib-32.dll" Alias "ibeot" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibfind32 Lib "Gpib-32.dll" Alias "ibfindA" (sstr As Any) As Long
Declare Function ibgts32 Lib "Gpib-32.dll" Alias "ibgts" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibist32 Lib "Gpib-32.dll" Alias "ibist" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function iblck32 Lib "Gpib-32.dll" Alias "iblck" (ByVal ud As Long, ByVal v As Long, ByVal LockWaitTime As Long, arg1 As Any) As Long
Declare Function iblines32 Lib "Gpib-32.dll" Alias "iblines" (ByVal ud As Long, v As Long) As Long
Declare Function ibln32 Lib "Gpib-32.dll" Alias "ibln" (ByVal ud As Long, ByVal pad As Long, ByVal sad As Long, ln As Long) As Long
Declare Function ibloc32 Lib "Gpib-32.dll" Alias "ibloc" (ByVal ud As Long) As Long
Declare Function iblock32 Lib "Gpib-32.dll" Alias "iblock" (ByVal ud As Long) As Long
Declare Function ibnotify32 Lib "Gpib-32.dll" Alias "ibnotify" (ByVal ud As Long, ByVal mask As Long, ByVal Callback As Long, RefData As Any) As Long
Declare Function ibonl32 Lib "Gpib-32.dll" Alias "ibonl" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibpad32 Lib "Gpib-32.dll" Alias "ibpad" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibpct32 Lib "Gpib-32.dll" Alias "ibpct" (ByVal ud As Long) As Long
Declare Function ibppc32 Lib "Gpib-32.dll" Alias "ibppc" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibrd32 Lib "Gpib-32.dll" Alias "ibrd" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibrda32 Lib "Gpib-32.dll" Alias "ibrda" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibrdf32 Lib "Gpib-32.dll" Alias "ibrdfA" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibrpp32 Lib "Gpib-32.dll" Alias "ibrpp" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibrsc32 Lib "Gpib-32.dll" Alias "ibrsc" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibrsp32 Lib "Gpib-32.dll" Alias "ibrsp" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibrsv32 Lib "Gpib-32.dll" Alias "ibrsv" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibsad32 Lib "Gpib-32.dll" Alias "ibsad" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibsic32 Lib "Gpib-32.dll" Alias "ibsic" (ByVal ud As Long) As Long
Declare Function ibsre32 Lib "Gpib-32.dll" Alias "ibsre" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibstop32 Lib "Gpib-32.dll" Alias "ibstop" (ByVal ud As Long) As Long
Declare Function ibtmo32 Lib "Gpib-32.dll" Alias "ibtmo" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibtrg32 Lib "Gpib-32.dll" Alias "ibtrg" (ByVal ud As Long) As Long
Declare Function ibunlock32 Lib "Gpib-32.dll" Alias "ibunlock" (ByVal ud As Long) As Long
Declare Function ibwait32 Lib "Gpib-32.dll" Alias "ibwait" (ByVal ud As Long, ByVal mask As Long) As Long
Declare Function ibwrt32 Lib "Gpib-32.dll" Alias "ibwrt" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibwrta32 Lib "Gpib-32.dll" Alias "ibwrta" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibwrtf32 Lib "Gpib-32.dll" Alias "ibwrtfA" (ByVal ud As Long, sstr As Any) As Long
Declare Sub AllSpoll32 Lib "Gpib-32.dll" Alias "AllSpoll" (ByVal boardID As Long, arg1 As Any, arg2 As Any)
Declare Sub DevClear32 Lib "Gpib-32.dll" Alias "DevClear" (ByVal boardID As Long, ByVal v As Long)
Declare Sub DevClearList32 Lib "Gpib-32.dll" Alias "DevClearList" (ByVal boardID As Long, arg1 As Any)
Declare Sub EnableLocal32 Lib "Gpib-32.dll" Alias "EnableLocal" (ByVal boardID As Long, arg1 As Any)
Declare Sub EnableRemote32 Lib "Gpib-32.dll" Alias "EnableRemote" (ByVal boardID As Long, arg1 As Any)
Declare Sub FindLstn32 Lib "Gpib-32.dll" Alias "FindLstn" (ByVal boardID As Long, arg1 As Any, arg2 As Any, ByVal limit As Long)
Declare Sub FindRQS32 Lib "Gpib-32.dll" Alias "FindRQS" (ByVal boardID As Long, arg1 As Any, result As Long)
Declare Sub PassControl32 Lib "Gpib-32.dll" Alias "PassControl" (ByVal boardID As Long, ByVal addr As Long)
Declare Sub PPoll32 Lib "Gpib-32.dll" Alias "PPoll" (ByVal boardID As Long, result As Long)
Declare Sub PPollConfig32 Lib "Gpib-32.dll" Alias "PPollConfig" (ByVal boardID As Long, ByVal addr As Long, ByVal line As Long, ByVal sense As Long)
Declare Sub PPollUnconfig32 Lib "Gpib-32.dll" Alias "PPollUnconfig" (ByVal boardID As Long, arg1 As Any)
Declare Sub RcvRespMsg32 Lib "Gpib-32.dll" Alias "RcvRespMsg" (ByVal boardID As Long, arg1 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub ReadStatusByte32 Lib "Gpib-32.dll" Alias "ReadStatusByte" (ByVal boardID As Long, ByVal addr As Long, result As Long)
Declare Sub Receive32 Lib "Gpib-32.dll" Alias "Receive" (ByVal boardID As Long, ByVal addr As Long, arg1 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub ReceiveSetup32 Lib "Gpib-32.dll" Alias "ReceiveSetup" (ByVal boardID As Long, ByVal add As Long)
Declare Sub ResetSys32 Lib "Gpib-32.dll" Alias "ResetSys" (ByVal boardID As Long, arg1 As Any)
Declare Sub Send32 Lib "Gpib-32.dll" Alias "Send" (ByVal boardID As Long, ByVal addr As Long, sstr As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendCmds32 Lib "Gpib-32.dll" Alias "SendCmds" (ByVal boardID As Long, sstr As Any, ByVal cnt As Long)
Declare Sub SendDataBytes32 Lib "Gpib-32.dll" Alias "SendDataBytes" (ByVal boardID As Long, sstr As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendIFC32 Lib "Gpib-32.dll" Alias "SendIFC" (ByVal boardID As Long)
Declare Sub SendList32 Lib "Gpib-32.dll" Alias "SendList" (ByVal boardID As Long, arg1 As Any, arg2 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendLLO32 Lib "Gpib-32.dll" Alias "SendLLO" (ByVal boardID As Long)
Declare Sub SendSetup32 Lib "Gpib-32.dll" Alias "SendSetup" (ByVal boardID As Long, arg1 As Any)
Declare Sub SetRWLS32 Lib "Gpib-32.dll" Alias "SetRWLS" (ByVal boardID As Long, arg1 As Any)
Declare Sub TestSys32 Lib "Gpib-32.dll" Alias "TestSys" (ByVal boardID As Long, arg1 As Any, arg2 As Any)
Declare Sub Trigger32 Lib "Gpib-32.dll" Alias "Trigger" (ByVal boardID As Long, ByVal addr As Long)
Declare Sub TriggerList32 Lib "Gpib-32.dll" Alias "TriggerList" (ByVal boardID As Long, arg1 As Any)
Declare Function RegisterGpibGlobalsForThread Lib "Gpib-32.dll" (Longibsta As Long, Longiberr As Long, Longibcnt As Long, ibcntl As Long) As Long
Declare Function UnregisterGpibGlobalsForThread Lib "Gpib-32.dll" () As Long
Declare Function ThreadIbsta32 Lib "Gpib-32.dll" Alias "ThreadIbsta" () As Long
Declare Function ThreadIbcnt32 Lib "Gpib-32.dll" Alias "ThreadIbcnt" () As Long
Declare Function ThreadIbcntl32 Lib "Gpib-32.dll" Alias "ThreadIbcntl" () As Long
Declare Function ThreadIberr32 Lib "Gpib-32.dll" Alias "ThreadIberr" () As Long
Declare Function iblockx32 Lib "Gpib-32.dll" Alias "iblockxA" (ByVal ud As Long, ByVal LockWaitTime As Long, arg1 As Any) As Long
Declare Function ibunlockx32 Lib "Gpib-32.dll" Alias "ibunlockx" (ByVal ud As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub AllSpoll(ByVal boardID As Integer, addrs() As Integer, results() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call AllSpoll32(boardID, addrs(0), results(0))
   
   Call copy_ibvars
End Sub

Sub copy_ibvars()
   ibsta = ConvertLongToInt(Longibsta)
   iberr = CInt(Longiberr)
   ibcnt = ConvertLongToInt(ibcntl)
End Sub

Sub DevClear(ByVal boardID As Integer, ByVal addr As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call DevClear32(boardID, addr)
   
   Call copy_ibvars
End Sub

Sub DevClearList(ByVal boardID As Integer, addrs() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call DevClearList32(boardID, addrs(0))
   
   Call copy_ibvars
End Sub

Sub EnableLocal(ByVal boardID As Integer, addrs() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call EnableLocal32(boardID, addrs(0))
   
   Call copy_ibvars
End Sub

Sub EnableRemote(ByVal boardID As Integer, addrs() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call EnableRemote32(boardID, addrs(0))
   
   Call copy_ibvars
End Sub

Sub FindLstn(ByVal boardID As Integer, addrs() As Integer, results() As Integer, ByVal limit As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call FindLstn32(boardID, addrs(0), results(0), limit)
   
   Call copy_ibvars
End Sub

Sub FindRQS(ByVal boardID As Integer, addrs() As Integer, result As Integer)
   Dim tmpresult As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call FindRQS32(boardID, addrs(0), tmpresult)
   
   result = ConvertLongToInt(tmpresult)
   
   Call copy_ibvars
End Sub

Sub ibask(ByVal ud As Integer, ByVal opt As Integer, rval As Integer)
   Dim tmprval As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibask32(ud, opt, tmprval)
   
   rval = ConvertLongToInt(tmprval)
   
   Call copy_ibvars
End Sub

Sub ibbna(ByVal ud As Integer, ByVal udname As String)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibbna32(ud, ByVal udname)
   
   Call copy_ibvars
End Sub

Sub ibcac(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibcac32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibclr(ByVal ud As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibclr32(ud)
   
   Call copy_ibvars
End Sub

Sub ibcmd(ByVal ud As Integer, ByVal buf As String)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   Call ibcmd32(ud, ByVal buf, cnt)
   
   Call copy_ibvars
End Sub

Sub ibcmda(ByVal ud As Integer, ByVal buf As String)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   ' synchronous call
   Call ibcmd32(ud, ByVal buf, cnt)
   
   Call copy_ibvars
End Sub

Sub ibconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibconfig32(bdid, opt, v)
   
   Call copy_ibvars
End Sub

Sub ibdev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer, ud As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ud = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

   Call copy_ibvars
End Sub

Sub ibdma(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibdma32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibeos(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibeos32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibeot(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibeot32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibfind(ByVal udname As String, ud As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ud = ConvertLongToInt(ibfind32(ByVal udname))
   
   Call copy_ibvars
End Sub

Sub ibgts(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibgts32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibist(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibist32(ud, v)
   
   Call copy_ibvars
End Sub

Sub iblines(ByVal ud As Integer, lines As Integer)
   Dim tmplines As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call iblines32(ud, tmplines)
   
   lines = ConvertLongToInt(tmplines)
   
   Call copy_ibvars
End Sub

Sub ibln(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer)
   Dim tmpln As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibln32(ud, pad, sad, tmpln)
   
   ln = ConvertLongToInt(tmpln)
   
   Call copy_ibvars
End Sub

Sub ibloc(ByVal ud As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibloc32(ud)
   
   Call copy_ibvars
End Sub

Sub iblck(ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Long)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call iblck32(ud, v, LockWaitTime, ByVal 0)
   
   Call copy_ibvars
  
End Sub

Sub ibnotify(ByVal ud As Integer, ByVal mask As Integer, Callback As Long, ByRef RefData As Variant)
    ' Register Globals if necessary
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    Call ibnotify32(ud, mask, Callback, RefData)
    
    Call copy_ibvars
End Sub

Sub ibonl(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibonl32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibpad(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibpad32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibpct(ByVal ud As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibpct32(ud)
   
   Call copy_ibvars
End Sub

Sub ibppc(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibppc32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibrd(ByVal ud As Integer, buf As String)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   Call ibrd32(ud, ByVal buf, cnt)
   
   Call copy_ibvars
End Sub

Sub ibrda(ByVal ud As Integer, buf As String)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))

   ' synchronous call
   Call ibrd32(ud, ByVal buf, cnt)
   
   Call copy_ibvars
End Sub

Sub ibrdf(ByVal ud As Integer, ByVal filename As String)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibrdf32(ud, ByVal filename)
   
   Call copy_ibvars
End Sub

Sub ibrdi(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibrd32(ud, ibuf(0), cnt)
   
   Call copy_ibvars
End Sub

Sub ibrdia(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ' synchronous call
   Call ibrd32(ud, ibuf(0), cnt)
   
   Call copy_ibvars
End Sub

Sub ibrpp(ByVal ud As Integer, ppr As Integer)
   Static tmp_str As String * 2
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibrpp32(ud, ByVal tmp_str)
   
   ppr = Asc(tmp_str)
   
   Call copy_ibvars
End Sub

Sub ibrsc(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibrsc32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibrsp(ByVal ud As Integer, spr As Integer)
   Static tmp_str As String * 2
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibrsp32(ud, ByVal tmp_str)
   
   spr = Asc(tmp_str)
   
   Call copy_ibvars
End Sub

Sub ibrsv(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibrsv32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibsad(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibsad32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibsic(ByVal ud As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibsic32(ud)
   
   Call copy_ibvars
End Sub

Sub ibsre(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibsre32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibstop(ByVal ud As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibstop32(ud)
   
   Call copy_ibvars
End Sub

Sub ibtmo(ByVal ud As Integer, ByVal v As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibtmo32(ud, v)
   
   Call copy_ibvars
End Sub

Sub ibtrg(ByVal ud As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibtrg32(ud)
   
   Call copy_ibvars
End Sub

Sub ibwait(ByVal ud As Integer, ByVal mask As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibwait32(ud, mask)
   
   Call copy_ibvars
End Sub

Sub ibwrt(ByVal ud As Integer, ByVal buf As String)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   Call ibwrt32(ud, ByVal buf, cnt)
   
   Call copy_ibvars
End Sub

Sub ibwrta(ByVal ud As Integer, ByVal buf As String)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   ' synchronous call
   Call ibwrt32(ud, ByVal buf, cnt)
   
   Call copy_ibvars
End Sub

Sub ibwrtf(ByVal ud As Integer, ByVal filename As String)

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibwrtf32(ud, ByVal filename)
   
   Call copy_ibvars
End Sub

Sub ibwrti(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibwrt32(ud, ibuf(0), cnt)
   
   Call copy_ibvars
End Sub

Sub ibwrtia(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ' synchronous call
   Call ibwrt32(ud, ibuf(0), cnt)
   
   Call copy_ibvars
End Sub

Function ilask(ByVal ud As Integer, ByVal opt As Integer, rval As Integer) As Integer
   Dim tmprval As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilask = ConvertLongToInt(ibask32(ud, opt, tmprval))
   
   rval = ConvertLongToInt(tmprval)
   
   Call copy_ibvars
End Function

Function ilbna(ByVal ud As Integer, ByVal udname As String) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilbna = ConvertLongToInt(ibbna32(ud, ByVal udname))
   
   Call copy_ibvars
End Function

Function ilcac(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilcac = ConvertLongToInt(ibcac32(ud, v))
   
   Call copy_ibvars
End Function

Function ilclr(ByVal ud As Integer) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilclr = ConvertLongToInt(ibclr32(ud))
   
   Call copy_ibvars
End Function

Function ilcmd(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilcmd = ConvertLongToInt(ibcmd32(ud, ByVal buf, cnt))
   
   Call copy_ibvars
End Function

Function ilcmda(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ' Synchronous call
   ilcmda = ConvertLongToInt(ibcmd32(ud, ByVal buf, cnt))
   
   Call copy_ibvars
End Function

Function ilconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal v As Integer) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilconfig = ConvertLongToInt(ibconfig32(bdid, opt, v))
   
   Call copy_ibvars
End Function

Function ildev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ildev = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))
   
   Call copy_ibvars
End Function

Function ildma(ByVal ud As Integer, ByVal v As Integer) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ildma = ConvertLongToInt(ibdma32(ud, v))
   
   Call copy_ibvars
End Function

Function ileos(ByVal ud As Integer, ByVal v As Integer) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ileos = ConvertLongToInt(ibeos32(ud, v))
   
   Call copy_ibvars
End Function

Function ileot(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ileot = ConvertLongToInt(ibeot32(ud, v))
   
   Call copy_ibvars
End Function

Function ilfind(ByVal udname As String) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilfind = ConvertLongToInt(ibfind32(ByVal udname))
   
   Call copy_ibvars
End Function

Function ilgts(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilgts = ConvertLongToInt(ibgts32(ud, v))
   
   Call copy_ibvars
End Function

Function ilist(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilist = ConvertLongToInt(ibist32(ud, v))
   
   Call copy_ibvars
End Function

Function illck(ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Long) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   illck = ConvertLongToInt(iblck32(ud, v, LockWaitTime, ByVal 0))
   
   Call copy_ibvars
End Function

Function illines(ByVal ud As Integer, lines As Integer) As Integer
   Dim tmplines As Long
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   illines = ConvertLongToInt(iblines32(ud, tmplines))
   
   lines = ConvertLongToInt(tmplines)
   
   Call copy_ibvars
End Function

Function illn(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer) As Integer
   Dim tmpln As Long
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   illn = ConvertLongToInt(ibln32(ud, pad, sad, tmpln))
   
   ln = ConvertLongToInt(tmpln)
   
   Call copy_ibvars
End Function

Function illoc(ByVal ud As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   illoc = ConvertLongToInt(ibloc32(ud))
   
   Call copy_ibvars
End Function

Function ilonl(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilonl = ConvertLongToInt(ibonl32(ud, v))
   
   Call copy_ibvars
End Function

Function ilnotify(ByVal ud As Integer, ByVal mask As Integer, Callback As Long, ByRef RefData As Variant) As Integer
   ' Register Globals if necessary
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    ilnotify = ConvertLongToInt(ibnotify32(ud, mask, Callback, RefData))
    
    Call copy_ibvars
End Function

Function ilpad(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilpad = ConvertLongToInt(ibpad32(ud, v))
   
   Call copy_ibvars
End Function

Function ilpct(ByVal ud As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilpct = ConvertLongToInt(ibpct32(ud))
   
   Call copy_ibvars
End Function

Function ilppc(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilppc = ConvertLongToInt(ibppc32(ud, v))
   
   Call copy_ibvars
End Function

Function ilrd(ByVal ud As Integer, buf As String, ByVal cnt As Long) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilrd = ConvertLongToInt(ibrd32(ud, ByVal buf, cnt))
   
   Call copy_ibvars
End Function

Function ilrda(ByVal ud As Integer, buf As String, ByVal cnt As Long) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ' synchronous call
   ilrda = ConvertLongToInt(ibrd32(ud, ByVal buf, cnt))
   
   Call copy_ibvars
End Function

Function ilrdf(ByVal ud As Integer, ByVal filename As String) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilrdf = ConvertLongToInt(ibrdf32(ud, ByVal filename))
   
   Call copy_ibvars
End Function

Function ilrdi(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilrdi = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))
   
   Call copy_ibvars
End Function

Function ilrdia(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ' synchronous call
   ilrdia = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))
   
   Call copy_ibvars
End Function

Function ilrpp(ByVal ud As Integer, ppr As Integer) As Integer
   Static tmpstr As String * 2
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilrpp = ConvertLongToInt(ibrpp32(ud, ByVal tmpstr))
   
   ppr = Asc(tmpstr)
   
   Call copy_ibvars
End Function

Function ilrsc(ByVal ud As Integer, ByVal v As Integer) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilrsc = ConvertLongToInt(ibrsc32(ud, v))
   
   Call copy_ibvars
End Function

Function ilrsp(ByVal ud As Integer, spr As Integer) As Integer
   Static tmpstr As String * 2
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilrsp = ConvertLongToInt(ibrsp32(ud, ByVal tmpstr))
   
   spr = Asc(tmpstr)
   
   Call copy_ibvars
End Function

Function ilrsv(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilrsv = ConvertLongToInt(ibrsv32(ud, v))
   
   Call copy_ibvars
End Function

Function ilsad(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilsad = ConvertLongToInt(ibsad32(ud, v))
   
   Call copy_ibvars
End Function

Function ilsic(ByVal ud As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilsic = ConvertLongToInt(ibsic32(ud))
   
   Call copy_ibvars
End Function

Function ilsre(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilsre = ConvertLongToInt(ibsre32(ud, v))
   
   Call copy_ibvars
End Function

Function ilstop(ByVal ud As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilstop = ConvertLongToInt(ibstop32(ud))
   
   Call copy_ibvars
End Function

Function iltmo(ByVal ud As Integer, ByVal v As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   iltmo = ConvertLongToInt(ibtmo32(ud, v))
   
   Call copy_ibvars
End Function

Function iltrg(ByVal ud As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   iltrg = ConvertLongToInt(ibtrg32(ud))
   
   Call copy_ibvars
End Function

Function ilwait(ByVal ud As Integer, ByVal mask As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilwait = ConvertLongToInt(ibwait32(ud, mask))
   
   Call copy_ibvars
End Function

Function ilwrt(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilwrt = ConvertLongToInt(ibwrt32(ud, ByVal buf, cnt))
   
   Call copy_ibvars
End Function

Function ilwrta(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If

   'synchronous call
   ilwrta = ConvertLongToInt(ibwrt32(ud, ByVal buf, cnt))
   
   Call copy_ibvars
End Function

Function ilwrtf(ByVal ud As Integer, ByVal filename As String) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilwrtf = ConvertLongToInt(ibwrtf32(ud, ByVal filename))
   
   Call copy_ibvars
End Function

Function ilwrti(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilwrti = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))
   
   Call copy_ibvars
End Function

Function ilwrtia(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   'synchronous call
   ilwrtia = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))
   
   Call copy_ibvars
End Function

Sub PassControl(ByVal boardID As Integer, ByVal addr As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call PassControl32(boardID, addr)
   
   Call copy_ibvars
End Sub

Sub Ppoll(ByVal boardID As Integer, result As Integer)
   Dim tmpresult As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call PPoll32(boardID, tmpresult)
   
   result = ConvertLongToInt(tmpresult)
   
   Call copy_ibvars
End Sub

Sub PpollConfig(ByVal boardID As Integer, ByVal addr As Integer, ByVal lline As Integer, ByVal sense As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call PPollConfig32(boardID, addr, lline, sense)
   
   Call copy_ibvars
End Sub

Sub PpollUnconfig(ByVal boardID As Integer, addrs() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call PPollUnconfig32(boardID, addrs(0))
   
   Call copy_ibvars
End Sub

Sub RcvRespMsg(ByVal boardID As Integer, buf As String, ByVal term As Integer)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   Call RcvRespMsg32(boardID, ByVal buf, cnt, term)
   
   Call copy_ibvars
End Sub

Sub ReadStatusByte(ByVal boardID As Integer, ByVal addr As Integer, result As Integer)
   Dim tmpresult As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ReadStatusByte32(boardID, addr, tmpresult)
   
   result = ConvertLongToInt(tmpresult)
   
   Call copy_ibvars
End Sub

Sub Receive(ByVal boardID As Integer, ByVal addr As Integer, buf As String, ByVal term As Integer)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   Call Receive32(boardID, addr, ByVal buf, cnt, term)
   
   Call copy_ibvars
End Sub

Sub ReceiveSetup(ByVal boardID As Integer, ByVal addr As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ReceiveSetup32(boardID, addr)
   
   Call copy_ibvars
End Sub

Sub ResetSys(ByVal boardID As Integer, addrs() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ResetSys32(boardID, addrs(0))
   
   Call copy_ibvars
End Sub

Sub Send(ByVal boardID As Integer, ByVal addr As Integer, ByVal buf As String, ByVal term As Integer)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   Call Send32(boardID, addr, ByVal buf, cnt, term)
   
   Call copy_ibvars
End Sub

Sub SendCmds(ByVal boardID As Integer, ByVal cmdbuf As String)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(cmdbuf))
   
   Call SendCmds32(boardID, ByVal cmdbuf, cnt)
   
   Call copy_ibvars
End Sub

Sub SendDataBytes(ByVal boardID As Integer, ByVal buf As String, ByVal term As Integer)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   Call SendDataBytes32(boardID, ByVal buf, cnt, term)
   
   Call copy_ibvars
End Sub
 
Sub SendIFC(ByVal boardID As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call SendIFC32(boardID)
   
   Call copy_ibvars
End Sub

Sub SendList(ByVal boardID As Integer, addr() As Integer, ByVal buf As String, ByVal term As Integer)
   Dim cnt As Long
   
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   cnt = CLng(Len(buf))
   
   Call SendList32(boardID, addr(0), ByVal buf, cnt, term)
   
   Call copy_ibvars
End Sub

Sub SendLLO(ByVal boardID As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call SendLLO32(boardID)
   
   Call copy_ibvars
End Sub

Sub SendSetup(ByVal boardID As Integer, addrs() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call SendSetup32(boardID, addrs(0))
   
   Call copy_ibvars
End Sub

Sub SetRWLS(ByVal boardID As Integer, addrs() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call SetRWLS32(boardID, addrs(0))
   
   Call copy_ibvars
End Sub

Sub TestSRQ(ByVal boardID As Integer, result As Integer)
   Call ibwait(boardID, 0)
   
   If ibsta And &H1000 Then
      result = 1
   Else
      result = 0
   End If
   
End Sub

Sub TestSys(ByVal boardID As Integer, addrs() As Integer, results() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call TestSys32(boardID, addrs(0), results(0))
   
   Call copy_ibvars
End Sub

Sub Trigger(ByVal boardID As Integer, ByVal addr As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call Trigger32(boardID, addr)
   
   Call copy_ibvars
End Sub

Sub TriggerList(ByVal boardID As Integer, addrs() As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call TriggerList32(boardID, addrs(0))
   
   Call copy_ibvars
End Sub

Sub WaitSRQ(ByVal boardID As Integer, result As Integer)
   Call ibwait(boardID, &H5000)
   
   If ibsta And &H1000 Then
      result = 1
   Else
      result = 0
   End If
End Sub

Private Function ConvertLongToInt(LongNumb As Long) As Integer

   If (LongNumb And &H8000&) = 0 Then
      ConvertLongToInt = LongNumb And &HFFFF&
   Else
      ConvertLongToInt = &H8000 Or (LongNumb And &H7FFF&)
   End If
End Function

Public Sub RegisterGPIBGlobals()
   Dim rc As Long
   
   rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
   If (rc = 0) Then
      GPIBglobalsRegistered = 1
   ElseIf (rc = 1) Then
      rc = UnregisterGpibGlobalsForThread
      rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
      GPIBglobalsRegistered = 1
   ElseIf (rc = 2) Then
      rc = UnregisterGpibGlobalsForThread
      ibsta = &H8000
      iberr = EDVR
      ibcntl = &HDEAD37F0
   ElseIf (rc = 3) Then
      rc = UnregisterGpibGlobalsForThread
      ibsta = &H8000
      iberr = EDVR
      ibcntl = &HDEAD37F0
   Else
      ibsta = &H8000
      iberr = EDVR
      ibcntl = &HDEAD37F0
   End If
End Sub

Public Sub UnregisterGPIBGlobals()
   Dim rc As Long
   
   rc = UnregisterGpibGlobalsForThread
   GPIBglobalsRegistered = 0
End Sub

Public Function ThreadIbsta() As Integer
   ThreadIbsta = ConvertLongToInt(ThreadIbsta32())
End Function

Public Function ThreadIberr() As Integer
   ThreadIberr = ConvertLongToInt(ThreadIberr32())
End Function

Public Function ThreadIbcnt() As Integer
   ThreadIbcnt = ConvertLongToInt(ThreadIbcnt32())
End Function

Public Function ThreadIbcntl() As Long
   ThreadIbcntl = ThreadIbcntl32()
End Function

Public Function illock(ByVal ud As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   illock = ConvertLongToInt(iblock32(ud))
   
   Call copy_ibvars
End Function

Public Function ilunlock(ByVal ud As Integer) As Integer

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilunlock = ConvertLongToInt(ibunlock32(ud))
   
   Call copy_ibvars
End Function

Public Sub iblock(ByVal ud As Integer)

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call iblock32(ud)
   
   Call copy_ibvars
End Sub

Public Sub ibunlock(ByVal ud As Integer)

   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibunlock32(ud)
   
   Call copy_ibvars
End Sub

Public Function illockx(ByVal ud As Integer, ByVal LockWaitTime As Integer, ByVal buf As String) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   illockx = ConvertLongToInt(iblockx32(ud, LockWaitTime, buf))
   
   Call copy_ibvars
End Function

Public Function ilunlockx(ByVal ud As Integer) As Integer
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   ilunlockx = ConvertLongToInt(ibunlockx32(ud))
   
   Call copy_ibvars
   
End Function

Public Sub iblockx(ByVal ud As Integer, ByVal LockWaitTime As Integer, ByVal buf As String)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call iblockx32(ud, LockWaitTime, buf)
   
   Call copy_ibvars
End Sub

Public Sub ibunlockx(ByVal ud As Integer)
   ' Register Globals if necessary
   If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
   End If
   
   Call ibunlockx32(ud)
   
   Call copy_ibvars
   
End Sub
   
   



