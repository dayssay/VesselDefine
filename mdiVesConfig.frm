VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm mdiVesConfig 
   BackColor       =   &H00808080&
   Caption         =   "Vessel Configuration (v2.0)"
   ClientHeight    =   7350
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   16080
   Icon            =   "mdiVesConfig.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   Begin MSComDlg.CommonDialog cmdDialog 
      Left            =   7680
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6120
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiVesConfig.frx":038A
            Key             =   "imgCopy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiVesConfig.frx":0924
            Key             =   "imgVessel"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiVesConfig.frx":0EBE
            Key             =   "imgSave"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiVesConfig.frx":1458
            Key             =   "imgNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiVesConfig.frx":19F2
            Key             =   "imgLoad"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiVesConfig.frx":1F8C
            Key             =   "imgOpen"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgOpen"
            Object.ToolTipText     =   "Open Vessel from def File"
            ImageKey        =   "imgOpen"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgNew"
            Object.ToolTipText     =   "New Vessel"
            ImageKey        =   "imgNew"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgCopy"
            Object.ToolTipText     =   "Copy Vessel"
            ImageKey        =   "imgCopy"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgLoad"
            Object.ToolTipText     =   "Load Vessel"
            ImageKey        =   "imgLoad"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "imgSave"
            Object.ToolTipText     =   "Save Vessel Info."
            ImageKey        =   "imgSave"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "imgVessel"
            Object.ToolTipText     =   "Vessel Explorer"
            ImageKey        =   "imgVessel"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuCopyVes 
         Caption         =   "Co&py"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVesEx 
         Caption         =   "&Vessel Explorer"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "mdiVesConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
  Call InitBasic
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Unload frmConfig
  Unload frmNewVessel
  Unload frmLoadVessel
  Unload frmCopyVessel
  Unload frmVesExplorer
  
  G_Host_Con.Close
End Sub

Private Sub mnuConfig_Click()
  frmConfig.Show
  frmConfig.SetFocus
  Me.Enabled = False
End Sub

Private Sub mnuCopyVes_Click()
  frmCopyVessel.Show
  frmCopyVessel.SetFocus
  Me.Enabled = False
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuLoad_Click()
  frmLoadVessel.Show
  frmLoadVessel.SetFocus
  Me.Enabled = False
End Sub

Private Sub mnuNew_Click()
  frmNewVessel.Show
  frmNewVessel.SetFocus
  Me.Enabled = False
End Sub

Private Sub MNUSAVE_CLICK()
  Call SetSave
End Sub

Private Sub mnuOpen_Click()
  Call defFileOpen
End Sub

Private Sub mnuVesEx_Click()
  frmVesExplorer.Show
  frmVesExplorer.SetFocus
End Sub

Private Sub SetSave()
  Dim Res As Double
  Dim sConf As String
  Dim sInput As String
  
  Res = MsgBox("최종확인 후 저장인 경우 '예', 임시저장인 경우 '아니오', 저장하지 않는 경우 '취소' 를 선택하세요!", vbYesNoCancel, "Save Option")
  
  If Res = vbCancel Then Exit Sub
  
  If Res = vbYes Then
    sConf = gVessel.sConfNm
    If sConf = "" Then sConf = "시스템"
    sInput = InputBox("최종확인자 성명을 입력하세요. (4자리 이하)", "Input Box", sConf)
    
    If Len(sInput) > 5 Or Len(sInput) < 1 Then MsgBox "문자길이는 1자리 이상 5자리 이하여야 합니다.": Exit Sub
    
    gVessel.sConfNm = sInput
  ElseIf Res = vbNo Then
    gVessel.sConfNm = ""
  End If
  
  Call SaveVslStructure
  
  
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "imgOpen"
      Call defFileOpen
    Case "imgNew"
      frmNewVessel.Show
      frmNewVessel.SetFocus
      Me.Enabled = False
    Case "imgLoad"
      frmLoadVessel.Show
      frmLoadVessel.SetFocus
      Me.Enabled = False
    Case "imgCopy"
      frmCopyVessel.Show
      frmCopyVessel.SetFocus
      Me.Enabled = False
    Case "imgSave"
      Call SetSave
    Case "imgVessel"
      frmVesExplorer.Show
      frmVesExplorer.SetFocus
  End Select
End Sub

Private Sub SetDef61v(sFileName As String)
  Dim FileNumber, Strsize%
  Dim i%, j%, k%
  Dim tStr$, tByte As Byte, tSingle As Single, tInt%, tLong&
  Dim sHeader$
  Dim sVCode$, sVName$, sCallSign$, sInmarsat$, sLloyd$
  Dim nLoa As Single, nLbp As Single, nWidth As Single, nDepth As Single, nTopHgt As Single, nAntHgt As Single, nBgLength As Single
  Dim iDMaxRows%, iHMaxRows%, iStarboardRows%, iPortRows%, iDMaxTiers%, iHMaxTiers, iTopTierIdx%, iNoHatchs%, iNoBays%, iHD%, iRow%, iTier%
  Dim sBayNo$, iHchNo%
  Dim iMinTopTierIdx%, iMaxTopTierIdx%
  
  FileNumber = FreeFile
  
  On Error GoTo err
  
  If cmdDialog.FileName = "" Then Exit Sub
  
  Open sFileName For Binary As FileNumber
  
  Strsize = 50
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sHeader = tStr
  
  Strsize = 4
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sVCode = NullTrim(tStr)
  
  Strsize = 20
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sVName = NullTrim(tStr)
  
  Strsize = 6
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sCallSign = NullTrim(tStr)
  
  Strsize = 9
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sInmarsat = NullTrim(tStr)
  
  Strsize = 7
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sLloyd = NullTrim(tStr)
  
  For i = 1 To 9
    Get #FileNumber, , tSingle
    
    Select Case i
      Case 1
        nLoa = tSingle
      Case 2
        nLbp = tSingle
      Case 3
        nWidth = tSingle
      Case 4
        nDepth = tSingle
      Case 5
        nTopHgt = tSingle
      Case 6
        nAntHgt = tSingle
    End Select
  Next i
  
  'Particulars
  Get #FileNumber, , tInt
  iDMaxRows = tInt
  
  Get #FileNumber, , tInt
  iHMaxRows = tInt
  
  Get #FileNumber, , tInt
  iStarboardRows = tInt
  
  Get #FileNumber, , tInt
  iPortRows = tInt
  
  Get #FileNumber, , tInt
  iDMaxTiers = tInt
  
  Get #FileNumber, , tInt
  iHMaxTiers = tInt
  
  Get #FileNumber, , tInt
  iNoHatchs = tInt
  
  Get #FileNumber, , tInt
  
  Get #FileNumber, , tInt
  iNoBays = tInt
  
  Get #FileNumber, , tInt
  
  For i = 1 To 10
    Get #FileNumber, , tSingle
  Next i
  
  Get #FileNumber, , tInt
  
  Get #FileNumber, , tSingle
  nBgLength = tSingle
  
  If iNoHatchs > 0 And iNoBays > 0 Then
    If Val(sVCode) > 0 Or Len(sVCode) <> 4 Then
      'Close #FileNumber
      MsgBox "Define 정보의 Vessel Code 값이 유효하지 않습니다. - Vessel Code 값을 변경하세요!"
      sVCode = "PICT"
    End If
    
    Erase gVessel.gHatch
    Erase gVessel.gBay
    
    gVessel.sVCode = sVCode
    gVessel.sVName = sVName
    gVessel.sCallSign = sCallSign
    gVessel.sInmarsat = sInmarsat
    gVessel.sLloyd = sLloyd
    gVessel.nLoa = nLoa
    gVessel.nLbp = nLbp
    gVessel.nWidth = nWidth
    gVessel.nDepth = nDepth
    gVessel.nTopHgt = nTopHgt
    gVessel.nAntHgt = nAntHgt
    gVessel.iDMaxRows = iDMaxRows
    gVessel.iHMaxRows = iHMaxRows
    gVessel.iNoHatchs = iNoHatchs
    gVessel.iNoBays = iNoBays
    gVessel.iBgNo = iNoHatchs - 2
    gVessel.nBgLength = nBgLength
    
    If gVessel.iDMaxRows >= gVessel.iHMaxRows Then
      gVessel.iMaxRows = gVessel.iDMaxRows
    Else
      gVessel.iMaxRows = gVessel.iHMaxRows
    End If
    
    ReDim Preserve gVessel.gHatch(1 To gVessel.iNoHatchs)
    ReDim Preserve gVessel.gBay(1 To gVessel.iNoBays)
  Else
    Close #FileNumber
    MsgBox "Define 정보가 유효하지 않습니다. 6.1 버전 여부를 확인하세요!"
    Exit Sub
  End If
  
  Debug.Print Seek(FileNumber)
  
  Strsize = 2
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  For i = 1 To 14
    Get #FileNumber, , tSingle
  Next i
  
  Strsize = 4
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  For i = 1 To 16
    Get #FileNumber, , tInt
  Next i
  
  For i = 1 To 2
    Strsize = 20
    tStr = String$(Strsize, " ")
    Get #FileNumber, , tStr
  Next i
  
  Strsize = 49
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  If Seek(FileNumber) < 383 Then
    Strsize = 1
    tStr = String$(Strsize, " ")
    Get #FileNumber, , tStr
  End If
  
  Debug.Print Seek(FileNumber)
  
  For i = 1 To 9
    Get #FileNumber, , tSingle
  Next i
  
  For i = 1 To 12
    Get #FileNumber, , tSingle
  Next i
  
  Debug.Print Seek(FileNumber)
  
  Strsize = 33
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  If Seek(FileNumber) < 501 Then
    Strsize = 1
    tStr = String$(Strsize, " ")
    Get #FileNumber, , tStr
  End If
  
  Debug.Print Seek(FileNumber)
  
  'cell 정보
  For k = 1 To 75
    If k <= gVessel.iNoBays Then
      gVessel.gBay(k).iNoFileOpenCells = 0
    End If
    
    For i = 1 To 20
      For j = 1 To 20
        Get #FileNumber, , tInt
        If k <= gVessel.iNoBays And tInt > 0 Then
            gVessel.gBay(k).iNoFileOpenCells = gVessel.gBay(k).iNoFileOpenCells + 1
            ReDim Preserve gVessel.gBay(k).gFileOpenCell(1 To gVessel.gBay(k).iNoFileOpenCells)
            
            gVessel.gBay(k).gFileOpenCell(gVessel.gBay(k).iNoFileOpenCells).iTier = i
            gVessel.gBay(k).gFileOpenCell(gVessel.gBay(k).iNoFileOpenCells).iRow = j
        End If
      Next j
    Next i
  Next k
  
  Debug.Print Seek(FileNumber)
  
  'bay 정보
  iMinTopTierIdx = 20: iMaxTopTierIdx% = 0
  For i = 1 To 75
    Strsize = 2
    tStr = String$(Strsize, " ")
    Get #FileNumber, , tStr
    sBayNo = NullTrim(tStr)
    
    Strsize = 4
    tStr = String$(Strsize, " ")
    Get #FileNumber, , tStr
    
    Get #FileNumber, , tInt
    iHchNo = tInt
    
    Get #FileNumber, , tInt
    
    Get #FileNumber, , tInt
    
    Get #FileNumber, , tInt
    
    Get #FileNumber, , tInt
    iTopTierIdx = tInt
    
    For j = 1 To 20
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 8
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 12
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 12
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 12
      Get #FileNumber, , tInt
    Next j
    
    'Structure 정보 설정
    iPortRows = (iStarboardRows - iHMaxRows - CInt((iDMaxRows - iHMaxRows) / 2))
    
    If i <= gVessel.iNoBays Then
      If iTopTierIdx < iMinTopTierIdx And iTopTierIdx > 0 Then
        iMinTopTierIdx = iTopTierIdx
      End If
      If iTopTierIdx > iMaxTopTierIdx% Then
        iMaxTopTierIdx% = iTopTierIdx
      End If
      
      If i = 1 Then
        'Hatch 정보 설정
        For j = 1 To gVessel.iNoHatchs
          gVessel.gHatch(j).iHch = j
          gVessel.gHatch(j).sHchNo = Format(j, "00")
          gVessel.gHatch(j).iDTier = gVessel.iDMaxTiers
          gVessel.gHatch(j).iHTier = 1
          gVessel.gHatch(j).iDFrRow = 1
          gVessel.gHatch(j).iDToRow = gVessel.iDMaxRows
          gVessel.gHatch(j).iHFrRow = 1
          gVessel.gHatch(j).iHToRow = gVessel.iHMaxRows
          
          gVessel.gHatch(i).bSaveFlag = True
        Next j
      End If
      
      'Bay 정보 설정
      gVessel.gBay(i).iBay = i
      gVessel.gBay(i).sBayNo = sBayNo
      gVessel.gBay(i).iHchNo = iHchNo
      If Val(sBayNo) Mod 2 = 0 Then
        gVessel.gBay(i).sSize = "4"
      Else
        gVessel.gBay(i).sSize = "2"
      End If
      
      gVessel.gBay(i).iTopTierIdx = iTopTierIdx
      gVessel.gBay(i).bSaveFlag = True
      
      'Row Info
      gVessel.gBay(i).iNoRows(0) = gVessel.iHMaxRows
      gVessel.gBay(i).iNoRows(1) = gVessel.iDMaxRows
      ReDim gVessel.gBay(i).gRow(0 To 1, 1 To gVessel.iMaxRows)
      
      For j = 1 To gVessel.iHMaxRows
        gVessel.gBay(i).gRow(0, j).iHD = 0
        gVessel.gBay(i).gRow(0, j).iRow = j
        
        gVessel.gBay(i).gRow(0, j).sNo = ""
      Next j
        
      For j = 1 To gVessel.iDMaxRows
        gVessel.gBay(i).gRow(1, j).iHD = 1
        gVessel.gBay(i).gRow(1, j).iRow = j
        
        gVessel.gBay(i).gRow(1, j).sNo = ""
      Next j
      
    End If
  Next i
  
  'max tier 지정
  gVessel.iDMaxTiers = iDMaxTiers - iMinTopTierIdx + 1
  gVessel.iHMaxTiers = iMinTopTierIdx - iHMaxTiers + (iMaxTopTierIdx - iMinTopTierIdx) / 2
  If gVessel.iDMaxTiers >= gVessel.iHMaxTiers Then
    gVessel.iMaxTiers = gVessel.iDMaxTiers
  Else
    gVessel.iMaxTiers = gVessel.iHMaxTiers
  End If
  
  For i = 1 To gVessel.iNoBays
    'Tier Info
    gVessel.gBay(i).iNoTiers(0) = gVessel.iHMaxTiers
    gVessel.gBay(i).iNoTiers(1) = gVessel.iDMaxTiers
    ReDim gVessel.gBay(i).gTier(0 To 1, 1 To gVessel.iMaxTiers)
    
    For j = 1 To gVessel.iHMaxTiers
      gVessel.gBay(i).gTier(0, j).iHD = 0
      gVessel.gBay(i).gTier(0, j).iTier = j
      
      gVessel.gBay(i).gTier(0, j).sNo = ""
    Next j
      
    For j = 1 To gVessel.iDMaxTiers
      gVessel.gBay(i).gTier(1, j).iHD = 1
      gVessel.gBay(i).gTier(1, j).iTier = j
      
      gVessel.gBay(i).gTier(1, j).sNo = ""
    Next j
      
    'Cell Info
    ReDim gVessel.gBay(i).gCell(0 To 1, 1 To gVessel.iMaxRows, 1 To gVessel.iMaxTiers)
    ReDim gVessel.gBay(i).gCellOrg(0 To 1, 1 To gVessel.iMaxRows, 1 To gVessel.iMaxTiers)
    
    With gVessel.gBay(i)
      
      For j = 1 To .iNoFileOpenCells
        If .gFileOpenCell(j).iTier <= gVessel.gBay(i).iTopTierIdx Then
          iHD = 0
          iTier = .gFileOpenCell(j).iTier - (10 - gVessel.iHMaxTiers) - (gVessel.gBay(i).iTopTierIdx - 10)
          
          iRow = .gFileOpenCell(j).iRow - (iStarboardRows - iHMaxRows)
        Else
          iHD = 1
          iTier = .gFileOpenCell(j).iTier - gVessel.gBay(i).iTopTierIdx
          
          iRow = .gFileOpenCell(j).iRow - iPortRows
        End If
        If iRow > 0 And iRow <= gVessel.iMaxRows And iTier > 0 And iTier <= gVessel.iMaxTiers Then
          .gCell(iHD, iRow, iTier).iHD = iHD
          .gCell(iHD, iRow, iTier).iRow = iRow
          .gCell(iHD, iRow, iTier).iTier = iTier
          .gCell(iHD, iRow, iTier).sSS = "Y"
          
          .gCellOrg(iHD, iRow, iTier).iHD = iHD
          .gCellOrg(iHD, iRow, iTier).iRow = iRow
          .gCellOrg(iHD, iRow, iTier).iTier = iTier
          .gCellOrg(iHD, iRow, iTier).sSS = "Y"
        End If
      Next j
      
    End With
  Next i
  
  Debug.Print Seek(FileNumber)
  
  'Hatch cover
  Dim nF As Single, nT As Single, bFlag As Boolean
  For i = 1 To 75
    bFlag = False: k = 0
    For j = 1 To 80
      Get #FileNumber, , tInt
      If i <= gVessel.iNoBays Then
        If tInt >= -3 And tInt <= 3 And tInt <> 0 Then
          If bFlag = False Then
            nF = (j - (iPortRows * 4) - 1) / 4
            
            If nF >= 0 Then
              k = k + 1
              
              gVessel.gBay(i).iNoCvr = k
              ReDim Preserve gVessel.gBay(i).gHchCvr(1 To k)
              gVessel.gBay(i).gHchCvr(k).iCvr = k
              gVessel.gBay(i).gHchCvr(k).sHType = "O"
              gVessel.gBay(i).gHchCvr(k).iBay = i
              gVessel.gBay(i).gHchCvr(k).nF = nF
              gVessel.gBay(i).gHchCvr(k).nT = nF + 0.25
              
              bFlag = True
            End If
            
          Else
            nT = (j - (iPortRows * 4)) / 4
            If nT > 0 Then gVessel.gBay(i).gHchCvr(k).nT = nT
          End If
          
        Else
          bFlag = False
        End If
        
      End If
    Next j
    
    Get #FileNumber, , tInt
  Next i
  
  Debug.Print Seek(FileNumber)
  
  For i = 1 To 75
    For k = 1 To 25
      Strsize = 3
      tStr = String$(Strsize, " ")
      Get #FileNumber, , tStr
    Next k
  Next i
  
  Debug.Print Seek(FileNumber)
  
  'ROW/TIER NO.
  For i = 1 To 75
    For j = 1 To 60
      Strsize = 2
      tStr = String$(Strsize, " ")
      Get #FileNumber, , tStr
      
      If i <= gVessel.iNoBays Then
        If NullTrim(tStr) <> "" Then
          If j < 21 Then
            iRow = j - (iStarboardRows - iHMaxRows)
            If iRow > 0 And iRow <= gVessel.iMaxRows Then gVessel.gBay(i).gRow(0, iRow).sNo = NullTrim(tStr)
            
          ElseIf j < 41 Then
            If (j - 20) <= gVessel.gBay(i).iTopTierIdx Then
              iHD = 0
              iTier = (j - 20) - (10 - gVessel.iHMaxTiers) - (gVessel.gBay(i).iTopTierIdx - 10)
            Else
              iHD = 1
              iTier = (j - 20) - gVessel.gBay(i).iTopTierIdx
            End If
            
            If iTier > 0 And iTier <= gVessel.iMaxTiers Then gVessel.gBay(i).gTier(iHD, iTier).sNo = NullTrim(tStr)
          Else
            iRow = (j - 40) - iPortRows
            If iRow > 0 And iRow <= gVessel.iMaxRows Then gVessel.gBay(i).gRow(1, iRow).sNo = NullTrim(tStr)
          End If
        End If
      End If
    Next j
  Next i
  
  Debug.Print Seek(FileNumber)
  
  'Stacking weight
  For i = 1 To 75
    For j = 1 To 20
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 20
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 20
      Get #FileNumber, , tInt
    Next j
    
    'deck weight
    For j = 1 To 20
      Get #FileNumber, , tInt
      If i <= gVessel.iNoBays Then
        If tInt > 0 Then
          iRow = j - iPortRows
          If iRow > 0 And iRow <= gVessel.iMaxRows Then
            If gVessel.gBay(i).gRow(1, iRow).sNo <> "" Then
                gVessel.gBay(i).gRow(1, iRow).nStkWgt = tInt
            End If
          End If
          
        End If
      End If
    Next j
    
    'hold weight
    For j = 1 To 20
      Get #FileNumber, , tInt
      If i <= gVessel.iNoBays Then
        If tInt > 0 Then
          iRow = j - (iStarboardRows - iHMaxRows)
          If iRow > 0 And iRow <= gVessel.iMaxRows Then
            If gVessel.gBay(i).gRow(0, iRow).sNo <> "" Then
              gVessel.gBay(i).gRow(0, iRow).nStkWgt = tInt
            End If
          End If
          
        End If
      End If
    Next j
    
    For j = 1 To 20
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 21
      Get #FileNumber, , tInt
    Next j
  Next i
  
  Debug.Print Seek(FileNumber)
  
  Close #FileNumber
  
  gVessel.bSaveFlag = True
  gVesFlag = True
  
  Call SetBasic
  
  Exit Sub
err:
  MsgBox "File Open Error!!"
  
  Close #FileNumber
  
  Call InitBasic
End Sub

Private Sub SetDef65v(sFileName As String)
  Dim FileNumber, Strsize%
  Dim i%, j%, k%
  Dim tStr$, tByte As Byte, tSingle As Single, tInt%, tLong&
  Dim sHeader$
  Dim sVCode$, sVName$, sCallSign$, sInmarsat$, sLloyd$
  Dim nLoa As Single, nLbp As Single, nWidth As Single, nDepth As Single, nTopHgt As Single, nAntHgt As Single, nBgLength As Single
  Dim iDMaxRows%, iHMaxRows%, iStarboardRows%, iPortRows%, iDMaxTiers%, iHMaxTiers, iTopTierIdx%, iNoHatchs%, iNoBays%, iHD%, iRow%, iTier%
  Dim sBayNo$, iHchNo%
  Dim iMinTopTierIdx%, iMaxTopTierIdx%
  
  FileNumber = FreeFile
  
  On Error GoTo err
  
  Open sFileName For Binary As FileNumber
  
  Strsize = 50
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sHeader = tStr
  
  Strsize = 4
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sVCode = NullTrim(tStr)
  
  Strsize = 30
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sVName = NullTrim(tStr)
  
  Strsize = 6
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sCallSign = NullTrim(tStr)
  
  Strsize = 9
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sInmarsat = NullTrim(tStr)
  
  Strsize = 7
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sLloyd = NullTrim(tStr)
  
  For i = 1 To 9
    Get #FileNumber, , tSingle
    
    Select Case i
      Case 1
        nLoa = tSingle
      Case 2
        nLbp = tSingle
      Case 3
        nWidth = tSingle
      Case 4
        nDepth = tSingle
      Case 5
        nTopHgt = tSingle
      Case 6
        nAntHgt = tSingle
    End Select
  Next i
  
  'Particulars
  Get #FileNumber, , tInt
  iDMaxRows = tInt
  
  Get #FileNumber, , tInt
  iHMaxRows = tInt
  
  Get #FileNumber, , tInt
  iStarboardRows = tInt
  
  Get #FileNumber, , tInt
  iPortRows = tInt
  
  Get #FileNumber, , tInt
  iDMaxTiers = tInt
  
  Get #FileNumber, , tInt
  iHMaxTiers = tInt
  
  Get #FileNumber, , tInt
  iNoHatchs = tInt
  
  Get #FileNumber, , tInt
  
  Get #FileNumber, , tInt
  iNoBays = tInt
  
  Get #FileNumber, , tInt
  
  For i = 1 To 10
    Get #FileNumber, , tSingle
  Next i
  
  Get #FileNumber, , tInt
  
  Get #FileNumber, , tSingle
  nBgLength = tSingle
  
  If iNoHatchs > 0 And iNoBays > 0 Then
    If Val(sVCode) > 0 Or Len(sVCode) <> 4 Then
      'Close #FileNumber
      MsgBox "Define 정보의 Vessel Code 값이 유효하지 않습니다. - Vessel Code 값을 변경하세요!"
      sVCode = "PICT"
    End If
    
    Erase gVessel.gHatch
    Erase gVessel.gBay
    
    gVessel.sVCode = sVCode
    gVessel.sVName = sVName
    gVessel.sCallSign = sCallSign
    gVessel.sInmarsat = sInmarsat
    gVessel.sLloyd = sLloyd
    gVessel.nLoa = nLoa
    gVessel.nLbp = nLbp
    gVessel.nWidth = nWidth
    gVessel.nDepth = nDepth
    gVessel.nTopHgt = nTopHgt
    gVessel.nAntHgt = nAntHgt
    gVessel.iDMaxRows = iDMaxRows
    gVessel.iHMaxRows = iHMaxRows
    gVessel.iNoHatchs = iNoHatchs
    gVessel.iNoBays = iNoBays
    gVessel.iBgNo = iNoHatchs - 2
    gVessel.nBgLength = nBgLength
    
    If gVessel.iDMaxRows >= gVessel.iHMaxRows Then
      gVessel.iMaxRows = gVessel.iDMaxRows
    Else
      gVessel.iMaxRows = gVessel.iHMaxRows
    End If
    
    ReDim Preserve gVessel.gHatch(1 To gVessel.iNoHatchs)
    ReDim Preserve gVessel.gBay(1 To gVessel.iNoBays)
  Else
    Close #FileNumber
    MsgBox "Define 정보가 유효하지 않습니다. 6.5 버전 여부를 확인하세요!"
    Exit Sub
  End If
  
  Debug.Print Seek(FileNumber)
  
  
  Strsize = 2
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  For i = 1 To 15
    Get #FileNumber, , tSingle
  Next i
  
  For i = 1 To 16
    Get #FileNumber, , tInt
  Next i
  
  For i = 1 To 2
    Strsize = 20
    tStr = String$(Strsize, " ")
    Get #FileNumber, , tStr
  Next i
  
  Get #FileNumber, , tInt
  
  Strsize = 2
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  Get #FileNumber, , tInt
  
  Get #FileNumber, , tInt
  
  
  Get #FileNumber, , tSingle
  
  
  For i = 1 To 7
    Get #FileNumber, , tInt
  Next i
  
  Strsize = 1
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  Strsize = 87
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  For i = 1 To 24
    Get #FileNumber, , tSingle
  Next i
  
  Debug.Print Seek(FileNumber)
  
  Strsize = 22
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  Debug.Print Seek(FileNumber)
  
  
  Get #FileNumber, , tByte
  
  Get #FileNumber, , tByte
  
  Get #FileNumber, , tInt
  
  For i = 1 To 6
    Get #FileNumber, , tSingle
  Next i
  
  Strsize = 10
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  For i = 1 To 3
    Get #FileNumber, , tSingle
  Next i
  
  Strsize = 28
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  For i = 1 To 5
    Get #FileNumber, , tInt
  Next i
  
  Strsize = 38
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  Debug.Print Seek(FileNumber)
  
  'cell 정보
  For k = 1 To 100
    If k <= gVessel.iNoBays Then
      gVessel.gBay(k).iNoFileOpenCells = 0
    End If
    
    For i = 1 To 26
      For j = 1 To 26
        Get #FileNumber, , tInt
        If k <= gVessel.iNoBays And tInt <> 0 Then
            gVessel.gBay(k).iNoFileOpenCells = gVessel.gBay(k).iNoFileOpenCells + 1
            ReDim Preserve gVessel.gBay(k).gFileOpenCell(1 To gVessel.gBay(k).iNoFileOpenCells)
            
            gVessel.gBay(k).gFileOpenCell(gVessel.gBay(k).iNoFileOpenCells).iTier = i
            gVessel.gBay(k).gFileOpenCell(gVessel.gBay(k).iNoFileOpenCells).iRow = j
        End If
      Next j
    Next i
  Next k
  
  Debug.Print Seek(FileNumber)
  
  'bay 정보
  iMinTopTierIdx = 20: iMaxTopTierIdx% = 0
  For i = 1 To 100
    Strsize = 3
    tStr = String$(Strsize, " ")
    Get #FileNumber, , tStr
    sBayNo = NullTrim(tStr)
    
    Strsize = 4
    tStr = String$(Strsize, " ")
    Get #FileNumber, , tStr
    
    Get #FileNumber, , tInt
    iHchNo = tInt
    
    Get #FileNumber, , tInt
    
    Get #FileNumber, , tSingle
    
    Get #FileNumber, , tSingle
    
    Get #FileNumber, , tInt
    iTopTierIdx = tInt
    
    For j = 1 To 26
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 8
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 16
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 16
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 16
      Get #FileNumber, , tInt
    Next j
    
    Get #FileNumber, , tByte
    
    Get #FileNumber, , tByte
    
    Get #FileNumber, , tInt
    
    'Get #FileNumber, , tByte
    
    'Structure 정보 설정
    iPortRows = (iStarboardRows - iHMaxRows - CInt((iDMaxRows - iHMaxRows) / 2))
    
    If i <= gVessel.iNoBays Then
      If iTopTierIdx < iMinTopTierIdx And iTopTierIdx > 0 Then
        iMinTopTierIdx = iTopTierIdx
      End If
      If iTopTierIdx > iMaxTopTierIdx% Then
        iMaxTopTierIdx% = iTopTierIdx
      End If
      
      If i = 1 Then
        'Hatch 정보 설정
        For j = 1 To gVessel.iNoHatchs
          gVessel.gHatch(j).iHch = j
          gVessel.gHatch(j).sHchNo = Format(j, "00")
          gVessel.gHatch(j).iDTier = gVessel.iDMaxTiers
          gVessel.gHatch(j).iHTier = 1
          gVessel.gHatch(j).iDFrRow = 1
          gVessel.gHatch(j).iDToRow = gVessel.iDMaxRows
          gVessel.gHatch(j).iHFrRow = 1
          gVessel.gHatch(j).iHToRow = gVessel.iHMaxRows
          
          gVessel.gHatch(i).bSaveFlag = True
        Next j
      End If
      
      'Bay 정보 설정
      gVessel.gBay(i).iBay = i
      gVessel.gBay(i).sBayNo = sBayNo
      gVessel.gBay(i).iHchNo = iHchNo
      If Val(sBayNo) Mod 2 = 0 Then
        gVessel.gBay(i).sSize = "4"
      Else
        gVessel.gBay(i).sSize = "2"
      End If
      
      gVessel.gBay(i).iTopTierIdx = iTopTierIdx
      gVessel.gBay(i).bSaveFlag = True
      
      'Row Info
      gVessel.gBay(i).iNoRows(0) = gVessel.iHMaxRows
      gVessel.gBay(i).iNoRows(1) = gVessel.iDMaxRows
      ReDim gVessel.gBay(i).gRow(0 To 1, 1 To gVessel.iMaxRows)
      
      For j = 1 To gVessel.iHMaxRows
        gVessel.gBay(i).gRow(0, j).iHD = 0
        gVessel.gBay(i).gRow(0, j).iRow = j
        
        gVessel.gBay(i).gRow(0, j).sNo = ""
      Next j
        
      For j = 1 To gVessel.iDMaxRows
        gVessel.gBay(i).gRow(1, j).iHD = 1
        gVessel.gBay(i).gRow(1, j).iRow = j
        
        gVessel.gBay(i).gRow(1, j).sNo = ""
      Next j
      
    End If
  Next i
  
  'max tier 지정
  gVessel.iDMaxTiers = iDMaxTiers - iMinTopTierIdx + 1
  gVessel.iHMaxTiers = iMinTopTierIdx - iHMaxTiers + (iMaxTopTierIdx - iMinTopTierIdx) / 2
  If gVessel.iDMaxTiers >= gVessel.iHMaxTiers Then
    gVessel.iMaxTiers = gVessel.iDMaxTiers
  Else
    gVessel.iMaxTiers = gVessel.iHMaxTiers
  End If
  
  For i = 1 To gVessel.iNoBays
    'Tier Info
    gVessel.gBay(i).iNoTiers(0) = gVessel.iHMaxTiers
    gVessel.gBay(i).iNoTiers(1) = gVessel.iDMaxTiers
    ReDim gVessel.gBay(i).gTier(0 To 1, 1 To gVessel.iMaxTiers)
    
    For j = 1 To gVessel.iHMaxTiers
      gVessel.gBay(i).gTier(0, j).iHD = 0
      gVessel.gBay(i).gTier(0, j).iTier = j
      
      gVessel.gBay(i).gTier(0, j).sNo = ""
    Next j
      
    For j = 1 To gVessel.iDMaxTiers
      gVessel.gBay(i).gTier(1, j).iHD = 1
      gVessel.gBay(i).gTier(1, j).iTier = j
      
      gVessel.gBay(i).gTier(1, j).sNo = ""
    Next j
      
    'Cell Info
    ReDim gVessel.gBay(i).gCell(0 To 1, 1 To gVessel.iMaxRows, 1 To gVessel.iMaxTiers)
    ReDim gVessel.gBay(i).gCellOrg(0 To 1, 1 To gVessel.iMaxRows, 1 To gVessel.iMaxTiers)
    
    With gVessel.gBay(i)
      
      For j = 1 To .iNoFileOpenCells
        If .gFileOpenCell(j).iTier <= gVessel.gBay(i).iTopTierIdx Then
          iHD = 0
          iTier = .gFileOpenCell(j).iTier - (13 - gVessel.iHMaxTiers) - (gVessel.gBay(i).iTopTierIdx - 13)
          
          iRow = .gFileOpenCell(j).iRow - (iStarboardRows - iHMaxRows)
        Else
          iHD = 1
          iTier = .gFileOpenCell(j).iTier - gVessel.gBay(i).iTopTierIdx
          
          iRow = .gFileOpenCell(j).iRow - iPortRows
        End If
        If iRow > 0 And iRow <= gVessel.iMaxRows And iTier > 0 And iTier <= gVessel.iMaxTiers Then
          .gCell(iHD, iRow, iTier).iHD = iHD
          .gCell(iHD, iRow, iTier).iRow = iRow
          .gCell(iHD, iRow, iTier).iTier = iTier
          .gCell(iHD, iRow, iTier).sSS = "Y"
          
          .gCellOrg(iHD, iRow, iTier).iHD = iHD
          .gCellOrg(iHD, iRow, iTier).iRow = iRow
          .gCellOrg(iHD, iRow, iTier).iTier = iTier
          .gCellOrg(iHD, iRow, iTier).sSS = "Y"
        End If
      Next j
      
    End With
  Next i
  
  Debug.Print Seek(FileNumber)
  
  Strsize = 1300
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  
  'Hatch cover
  Dim nF As Single, nT As Single, bFlag As Boolean
  For i = 1 To 100
    bFlag = False: k = 0
    For j = 1 To 104
      Get #FileNumber, , tInt
      If i <= gVessel.iNoBays Then
        If tInt >= -3 And tInt <= 3 And tInt <> 0 Then
          If bFlag = False Then
            nF = (j - (iPortRows * 4) - 1) / 4
            
            If nF >= 0 Then
              k = k + 1
              
              gVessel.gBay(i).iNoCvr = k
              ReDim Preserve gVessel.gBay(i).gHchCvr(1 To k)
              gVessel.gBay(i).gHchCvr(k).iCvr = k
              gVessel.gBay(i).gHchCvr(k).sHType = "O"
              gVessel.gBay(i).gHchCvr(k).iBay = i
              gVessel.gBay(i).gHchCvr(k).nF = nF
              gVessel.gBay(i).gHchCvr(k).nT = nF + 0.25
              
              bFlag = True
            End If
            
          Else
            nT = (j - (iPortRows * 4)) / 4
            If nT > 0 Then gVessel.gBay(i).gHchCvr(k).nT = nT
          End If
          
        Else
          bFlag = False
        End If
        
      End If
    Next j
    
    Get #FileNumber, , tInt
  Next i
  
  Debug.Print Seek(FileNumber)
  
  
  For i = 1 To 100
    For j = 1 To 2
      For k = 1 To 25
        Strsize = 3
        tStr = String$(Strsize, " ")
        Get #FileNumber, , tStr
        
      Next k
    Next j
  Next i
  
  Debug.Print Seek(FileNumber)
  
  'ROW/TIER NO.
  For i = 1 To 100
    For j = 1 To 78
      Strsize = 3
      tStr = String$(Strsize, " ")
      Get #FileNumber, , tStr
      
      If i <= gVessel.iNoBays Then
        If NullTrim(tStr) <> "" Then
          If j < 27 Then
            iRow = j - (iStarboardRows - iHMaxRows)
            If iRow > 0 And iRow <= gVessel.iHMaxRows Then gVessel.gBay(i).gRow(0, iRow).sNo = NullTrim(tStr)
            
          ElseIf j < 53 Then
            If (j - 26) <= gVessel.gBay(i).iTopTierIdx Then
              iHD = 0
              'iTier = (j - 26) - iHMaxTiers
              iTier = (j - 26) - (13 - gVessel.iHMaxTiers) - (gVessel.gBay(i).iTopTierIdx - 13)
            Else
              iHD = 1
              iTier = (j - 26) - gVessel.gBay(i).iTopTierIdx
            End If
            
            If iTier > 0 And iTier <= gVessel.iMaxTiers Then gVessel.gBay(i).gTier(iHD, iTier).sNo = NullTrim(tStr)
          Else
            iRow = (j - 52) - iPortRows
            If iRow > 0 And iRow <= gVessel.iDMaxRows Then gVessel.gBay(i).gRow(1, iRow).sNo = NullTrim(tStr)
          End If
        End If
      End If
    Next j
  Next i
  
  Debug.Print Seek(FileNumber)
  
  'Stacking weight
  For i = 1 To 100
    For j = 1 To 26
      Get #FileNumber, , tLong
    Next j
    
    For j = 1 To 26
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 26
      Get #FileNumber, , tInt
    Next j
    
    'deck weight
    For j = 1 To 26
      Get #FileNumber, , tInt
      If i <= gVessel.iNoBays Then
        If tInt > 0 Then
          iRow = j - iPortRows
          If iRow > 0 And iRow <= gVessel.iMaxRows Then
            If gVessel.gBay(i).gRow(1, iRow).sNo <> "" Then
                gVessel.gBay(i).gRow(1, iRow).nStkWgt = tInt
            End If
          End If
          
        End If
      End If
    Next j
    
    'hold weight
    For j = 1 To 26
      Get #FileNumber, , tInt
      If i <= gVessel.iNoBays Then
        If tInt > 0 Then
          iRow = j - (iStarboardRows - iHMaxRows)
          If iRow > 0 And iRow <= gVessel.iMaxRows Then
            If gVessel.gBay(i).gRow(0, iRow).sNo <> "" Then
              gVessel.gBay(i).gRow(0, iRow).nStkWgt = tInt
            End If
          End If
          
        End If
      End If
    Next j
    
    For j = 1 To 26
      Get #FileNumber, , tInt
    Next j
    
    For j = 1 To 27
      Get #FileNumber, , tInt
    Next j
  Next i
  
  Debug.Print Seek(FileNumber)
  
  Close #FileNumber
  
  gVessel.bSaveFlag = True
  gVesFlag = True
  
  Call SetBasic
  
  Exit Sub
err:
  MsgBox "File Open Error!!"
  
  Close #FileNumber
  
  Call InitBasic
End Sub

Private Sub defFileOpen()
  Dim FileNumber, Strsize%
  Dim tStr$
  Dim sHeader$
  
  FileNumber = FreeFile
  
  On Error GoTo err
  
  cmdDialog.FileName = ""
  cmdDialog.DialogTitle = "def 파일 열기"
  cmdDialog.Filter = "def 파일(*.def)|*.def"
  cmdDialog.Flags = 0
  cmdDialog.CancelError = False
  cmdDialog.ShowOpen
  
  If cmdDialog.FileName = "" Then Exit Sub
  
  Open cmdDialog.FileName For Binary As FileNumber
  
  Strsize = 50
  tStr = String$(Strsize, " ")
  Get #FileNumber, , tStr
  sHeader = tStr
  
  
  Close #FileNumber
  If InStr(sHeader, "6.1") > 0 Then
    Call SetDef61v(cmdDialog.FileName)
  ElseIf InStr(sHeader, "6.5") > 0 Then
    Call SetDef65v(cmdDialog.FileName)
  Else
    MsgBox "Define 정보가 유효하지 않습니다. 6.1 / 6.5 버전 여부를 확인하세요!"
    Exit Sub
  End If
  
  Exit Sub
err:
  MsgBox "File Open Error!!"
  Close #FileNumber
End Sub
