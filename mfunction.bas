Attribute VB_Name = "mFunction"
Option Explicit

Public Function NullTrim(ss)
  On Error Resume Next

  If IsNull(ss) Then
    NullTrim = ""
  Else
    NullTrim = Trim(ss)
  End If
End Function

Public Function NullTrim_Num(ss)
  On Error Resume Next
  
  If IsNull(ss) Then
    NullTrim_Num = "0"
  ElseIf Trim(ss) = "" Then
    NullTrim_Num = "0"
  Else
    NullTrim_Num = Trim(ss)
  End If
End Function

Public Sub SetBasic()
  mdiVesConfig.Caption = "Vessel Configuration (v2.0) - " & gVessel.sVCode
  
  mdiVesConfig.mnuSave.Enabled = True
  mdiVesConfig.tbrMain.Buttons(6).Enabled = True
  
  mdiVesConfig.mnuVesEx.Enabled = True
  mdiVesConfig.tbrMain.Buttons(8).Enabled = True
  
  Unload frmVesExplorer
  
  frmVesExplorer.Show
End Sub

Public Sub InitBasic()
  Call SetSvrConfig
  Call Connect_DB
  gVesFlag = False
  
  mdiVesConfig.Caption = "Vessel Configuration (v2.0)"
  
  mdiVesConfig.mnuSave.Enabled = False
  mdiVesConfig.tbrMain.Buttons(6).Enabled = False
  
  mdiVesConfig.mnuVesEx.Enabled = False
  mdiVesConfig.tbrMain.Buttons(8).Enabled = False
  
  Unload frmVesExplorer
End Sub

Public Sub PasteCellInfo(iBay%, iBayTo%, iHD%, Optional sMode$, Optional iMaxRows%, Optional iMaxTiers%)
  Dim i%, j%
  
  If sMode = "C" Then
    For i = 1 To iMaxRows
      For j = 1 To iMaxTiers
        gVessel.gBay(iBay).gCell(iHD, i, j).sSS = gVessel.gBay(iBay).gCellOrg(iHD, i, j).sSS
        gVessel.gBay(iBay).gCell(iHD, i, j).sGuide = gVessel.gBay(iBay).gCellOrg(iHD, i, j).sGuide
      Next j
    Next i
  ElseIf sMode = "O" Then
    For i = 1 To iMaxRows
      For j = 1 To iMaxTiers
        gVessel.gBay(iBay).gCellOrg(iHD, i, j).sSS = gVessel.gBay(iBay).gCell(iHD, i, j).sSS
        gVessel.gBay(iBay).gCellOrg(iHD, i, j).sGuide = gVessel.gBay(iBay).gCell(iHD, i, j).sGuide
      Next j
    Next i
  Else
    '±âº» copy
    For i = 1 To gVessel.gBay(iBay).iNoRows(iHD)
      For j = 1 To gVessel.gBay(iBay).iNoTiers(iHD)
        gVessel.gBay(iBayTo).gCell(iHD, i, j).sSS = gVessel.gBay(iBay).gCell(iHD, i, j).sSS
        gVessel.gBay(iBayTo).gCell(iHD, i, j).sGuide = gVessel.gBay(iBay).gCell(iHD, i, j).sGuide
      Next j
    Next i
    
  End If
End Sub

Public Sub PasteCover(iBay%, iBayTo%)
  Dim i%
  
  gVessel.gBay(iBayTo).iNoCvr = gVessel.gBay(iBay).iNoCvr
  Erase gVessel.gBay(iBayTo).gHchCvr
  If gVessel.gBay(iBayTo).iNoCvr <= 0 Then Exit Sub
  ReDim gVessel.gBay(iBayTo).gHchCvr(1 To gVessel.gBay(iBayTo).iNoCvr)
  
  For i = 1 To gVessel.gBay(iBay).iNoCvr
    gVessel.gBay(iBayTo).gHchCvr(i).iCvr = gVessel.gBay(iBay).gHchCvr(i).iCvr
    gVessel.gBay(iBayTo).gHchCvr(i).iBay = iBayTo
    gVessel.gBay(iBayTo).gHchCvr(i).nF = gVessel.gBay(iBay).gHchCvr(i).nF
    gVessel.gBay(iBayTo).gHchCvr(i).nT = gVessel.gBay(iBay).gHchCvr(i).nT
    gVessel.gBay(iBayTo).gHchCvr(i).sHType = gVessel.gBay(iBay).gHchCvr(i).sHType
  Next i
End Sub

Public Sub PasteNo(iBay%, iBayTo%, iHD%)
  Dim i%
  
  For i = 1 To gVessel.gBay(iBay).iNoRows(iHD)
    gVessel.gBay(iBayTo).gRow(iHD, i).sNo = gVessel.gBay(iBay).gRow(iHD, i).sNo
    gVessel.gBay(iBayTo).gRow(iHD, i).nStkWgt = gVessel.gBay(iBay).gRow(iHD, i).nStkWgt
  Next i
  
  For i = 1 To gVessel.gBay(iBay).iNoTiers(iHD)
    gVessel.gBay(iBayTo).gTier(iHD, i).sNo = gVessel.gBay(iBay).gTier(iHD, i).sNo
  Next i
End Sub
