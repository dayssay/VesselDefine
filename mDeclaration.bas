Attribute VB_Name = "mDeclaration"
Option Explicit

'ADO관련 정의
Public G_Host_Con As ADODB.Connection

'쿼리관련 변수 정의
Public gRs As New ADODB.Recordset
Public SQL As String

'Server Config
Public Type tSvrCfg
  SvrIp As String
  SvrId As String
  SvrPw As String
  SvrDb As String
End Type
Public gSvrCfg As tSvrCfg

Public Const cSvrConfigFile = "Configuration.opt"

'Vessel Config 변수
Private Type tRow
  iHD As Integer
  iRow As Integer
  sNo As String
  nStkWgt As Integer
  nRowClr As Single
  iL As Single
  iT As Single
  iW As Single
  iH As Single
  
  iL_Wgt As Single
  iT_Wgt As Single
  iW_Wgt As Single
  iH_Wgt As Single
End Type

Private Type tTier
  iHD As Integer
  iTier As Integer
  sNo As String
  iL As Single
  iT As Single
  iW As Single
  iH As Single
End Type

Private Type tCell
  iHD As Integer
  iRow As Integer
  iTier As Integer
  sSS As String
  sGuide As String
  iL As Single
  iT As Single
  iW As Single
  iH As Single
End Type

Private Type tFileOpenCell
  iRow As Integer
  iTier As Integer
End Type

Private Type tHchCover
  iCvr As Integer
  sHType As String
  nF As Single
  nT As Single
  iBay As Integer
End Type

Private Type tBay
  iBay As Integer
  sBayNo As String
  nLcgDeck As Single
  nLcgHold As Single
  sSize As String
  iHchNo As Integer
  
  gRow() As tRow        'Row
  iNoRows(0 To 1) As Integer

  gTier() As tTier      'Tier
  iNoTiers(0 To 1) As Integer

  gCell() As tCell      'Cell
  gCellOrg() As tCell      'Cell
  'iNoCells As Integer
  
  'File Open 시 디파인 셀 정보
  gFileOpenCell() As tFileOpenCell
  iNoFileOpenCells As Integer
   
  'File Open 시 toptier index 정보
  iTopTierIdx As Integer
  
  'Hatch Cover
  gHchCvr() As tHchCover    'Hatch Cover
  iNoCvr As Integer
  
  bSaveFlag As Boolean
End Type

Private Type tHatchInfo
  iHch As Integer
  sHchNo As String
  iDTier As Integer
  iDFrRow As Integer
  iDToRow As Integer
  iHTier As Integer
  iHFrRow As Integer
  iHToRow As Integer
  
  bSaveFlag As Boolean
End Type

Private Type tVessel
  sVCode As String
  sVName As String
  sCallSign As String
  sInmarsat As String
  sLloyd As String
  nLoa As Single
  nLbp As Single
  nWidth As Single
  nDepth As Single
  nTopHgt As Single
  nAntHgt As Single
  iDMaxRows As Integer
  iDMaxTiers As Integer
  iHMaxRows As Integer
  iHMaxTiers As Integer
  iBgNo As Integer
  nBgLength As Single
  
  iMaxRows As Integer
  iMaxTiers As Integer
  
  gHatch() As tHatchInfo        'Hatch
  iNoHatchs As Integer
  
  gBay() As tBay            'Bay
  iNoBays As Integer
  
  sConfNm As String
  
  bSaveFlag As Boolean
End Type

Public gVessel As tVessel
Public gVesFlag As Boolean
'===================================



