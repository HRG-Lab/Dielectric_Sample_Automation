Attribute VB_Name = "NIGLOBAL"
'
'
'  Keysight 488 Implementation, Globals for Visual Basic
'
'
'
'
'

Option Explicit

Global ibsta As Integer
Global iberr As Integer
Global ibcnt As Integer
Global ibcntl As Long

Global Longibsta As Long
Global Longiberr As Long
Global Longibcnt As Long
Global GPIBglobalsRegistered As Integer

Global buf As String

Global bytebuf() As Byte

Global Const UNL = &H3F  
Global Const UNT = &H5F
Global Const GTL = &H1
Global Const SDC = &H4
Global Const PPC = &H5
Global Const GGET = &H8
Global Const TCT = &H9
Global Const LLO = &H11
Global Const DCL = &H14
Global Const PPU = &H15
Global Const SPE = &H18
Global Const SPD = &H19
Global Const PPE = &H60
Global Const PPD = &H70

' status bits

Global Const EERR = &H8000
Global Const TIMO = &H4000
Global Const EEND = &H2000
Global Const SRQI = &H1000
Global Const RQS  = &H0800
Global Const CMPL = &H0100
Global Const LOK  = &H0080
Global Const RREM = &H0040
Global Const CIC  = &H0020
Global Const AATN = &H0010
Global Const TACS = &H0008
Global Const LACS = &H0004
Global Const DTAS = &H0002
Global Const DCAS = &H0001

' Error Codes
Global Const EDVR = 0
Global Const ECIC = 1
Global Const ENOL = 2
Global Const EADR = 3
Global Const EARG = 4
Global Const ESAC = 5
Global Const EABO = 6
Global Const ENEB = 7
Global Const EDMA = 8
Global Const EOIP = 9
Global Const ECAP = 11
Global Const EFSO = 12
Global Const EBUS = 14
Global Const ESTB = 15
Global Const ESRQ = 16
Global Const ETAB = 20
Global Const ELCK = 21
Global Const EARM = 22
Global Const EHDL = 23
Global Const WCFG = 24
Global Const ECFG = 25
Global Const EWIP = 26
Global Const ERST = 27

' EOS mode bits

Global Const BIN  = &H1000
Global Const XEOS = &H0800
Global Const REOS = &H0400


' Timeout values

Global Const TNONE = 0
Global Const T10us = 1
Global Const T30us = 2
Global Const T100us = 3
Global Const T300us = 4
Global Const T1ms = 5
Global Const T3ms = 6
Global Const T10ms = 7
Global Const T30ms = 8
Global Const T100ms = 9
Global Const T300ms = 10
Global Const T1s = 11
Global Const T3s = 12
Global Const T10s = 13
Global Const T30s = 14
Global Const T100s = 15
Global Const T300s = 16
Global Const T1000s = 17

'IBLN constants

Global Const ALL_SAD = -1
Global Const NO_SAD = 0

' ibconfig option constants

Global Const IbcPAD = &H1
Global Const IbcSAD = &H2
Global Const IbcTMO = &H3
Global Const IbcEOT = &H4
Global Const IbcPPC = &H5
Global Const IbcREADDR = &H6
Global Const IbcAUTOPOLL = &H7
Global Const IbcCICPROT = &H8
Global Const IbcIRQ = &H9
Global Const IbcSC = &HA
Global Const IbcSRE = &HB
Global Const IbcEOSrd = &HC
Global Const IbcEOSwrt = &HD
Global Const IbcEOScmp = &HE
Global Const IbcEOSchar = &HF
Global Const IbcPP2 = &H10
Global Const IbcTIMING = &H11
Global Const IbcDMA = &H12
Global Const IbcReadADjust = &H13
Global Const IbcWriteAdjust = &H14
Global Const IbcSendLLO = &H17
Global Const IbcSPollTime = &H18
Global Const IbcPPollTime = &H19
Global Const IbcEndBitIsNormal = &H1A
Global Const IbcUnAddr = &H1B
Global Const IbcSignalNumber = &H1C
Global Const IbcBlockIfLocked = &H1D
Global Const IbcHSCableLength = &H1F
Global Const IbcIst = &H20
Global Const IbcRsv = &H21
Global Const IbcLON = &H22

' ibask option constants

Global Const IbaPAD = &H1
Global Const IbaSAD = &H2
Global Const IbaTMO = &H3
Global Const IbaEOT = &H4
Global Const IbaPPC = &H5
Global Const IbaREADDR = &H6
Global Const IbaAUTOPOLL = &H7
Global Const IbaCICPROT = &H8
Global Const IbaIRQ = &H9
Global Const IbaSC = &HA
Global Const IbaSRE = &HB
Global Const IbaEOSrd = &HC
Global Const IbaEOSwrt = &HD
Global Const IbaEOScmp = &HE
Global Const IbaEOSchar = &HF
Global Const IbaPP2 = &H10
Global Const IbaTIMING = &H11
Global Const IbaDMA = &H12
Global Const IbaReadADjust = &H13
Global Const IbaWriteAdjust = &H14
Global Const IbaSendLLO = &H17
Global Const IbaSPollTime = &H18
Global Const IbaPPollTime = &H19
Global Const IbaEndBitIsNormal = &H1A
Global Const IbaUnAddr = &H1B
Global Const IbaSignalNumber = &H1C
Global Const IbaBlockIfLocked = &H1D
Global Const IbaHSCableLength = &H1F
Global Const IbaIst = &H20
Global Const IbaRsv = &H21
Global Const IbaLON = &H22
Global Const IbaBNA = &H200

' Send command constants

Global Const NULLend = &H0
Global Const NLend = &H1
Global Const DABend = &H2

' Receive command constant

Global Const STOPend = &H100

'iblines constants

Global Const ValidEOI  = &H0080
Global Const ValidATN  = &H0040
Global Const ValidSRQ  = &H0020
Global Const ValidREN  = &H0010
Global Const ValidIFC  = &H0008
Global Const ValidNRFD = &H0004
Global Const ValidNDAC = &H0002
Global Const ValidDAV  = &H0001
Global Const BusEOI    = &H8000
Global Const BusATN    = &H4000
Global Const BusSRQ    = &H2000
Global Const BusREN    = &H1000
Global Const BusIFC    = &H0800
Global Const BusNRFD   = &H0400
Global Const BusNDAC   = &H0200
Global Const BusDAV    = &H0100

' Address List Constant

Global Const NOADDR = &HFFFF

' Callback Rearm error constant

Global Const IBNOTIFY_REARM_FAILED = &HE00A003F

' LAN lock constants (iblockx/inunlockx)
Global Const TIMMEDIATE = -1
Global Const TINFINITE = -2
Global Const MAX_LOCKSHARENAME_LENGTH = 64


