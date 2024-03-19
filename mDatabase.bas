Attribute VB_Name = "mDatabase"
Option Explicit

Public Function Connect_DB() As Boolean
  Dim sConnStr$
  
  'DB Connection
  Set G_Host_Con = New ADODB.Connection
  G_Host_Con.ConnectionTimeout = 15
  G_Host_Con.CommandTimeout = 1000
  
  'DNS »ç¿ë
'  sConnStr = "PROVIDER=MSDASQL;dsn=" & gSvrCfg.SvrDb & ";uid=" & gSvrCfg.SvrId & ";pwd=" & gSvrCfg.SvrPw & ";database=" & gSvrCfg.SvrDb & ";"
  
  sConnStr = "Provider=OraOLEDB.Oracle.1;Persist Security Info=True;User ID=" & gSvrCfg.SvrId & ";Password=" & gSvrCfg.SvrPw & ";"
  sConnStr = sConnStr & "Data Source=(DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = " & gSvrCfg.SvrIp & ")(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = " & gSvrCfg.SvrDb & ")))"
  G_Host_Con.Open sConnStr
  
End Function

Public Sub GetVslStructure(sVslCd$)
  Dim tRs As ADODB.Recordset
  
  Screen.MousePointer = 11
  
  'Block Info
  SQL = "SELECT SVCODE,"
  SQL = SQL & "        SVNAME,"
  SQL = SQL & "        SCALLSIGN,"
  SQL = SQL & "        SINMARSAT,"
  SQL = SQL & "        SLLOYD,"
  SQL = SQL & "        NLOA,"
  SQL = SQL & "        NLBP,"
  SQL = SQL & "        NWIDTH,"
  SQL = SQL & "        NDEPTH,"
  SQL = SQL & "        NTOPHGT,"
  SQL = SQL & "        NANTHGT,"
  SQL = SQL & "        INOBAY,"
  SQL = SQL & "        INOHATCH,"
  SQL = SQL & "        IDMAXROW,"
  SQL = SQL & "        IHMAXROW,"
  SQL = SQL & "        IDMAXTIER,"
  SQL = SQL & "        IHMAXTIER,"
  SQL = SQL & "        IBGNO,"
  SQL = SQL & "        NBGLENGTH,"
  SQL = SQL & "        INS_ID"
  SQL = SQL & "   FROM TB_VD_VESSEL"
  SQL = SQL & "  WHERE SVCODE = '" & sVslCd & "'"
  Set gRs = G_Host_Con.Execute(SQL)
  
  If Not (gRs.BOF And gRs.EOF) Then
    Erase gVessel.gHatch
    
    gVessel.sVCode = NullTrim(gRs!sVCode)
    gVessel.sVName = NullTrim(gRs!sVName)
    gVessel.sCallSign = NullTrim(gRs!sCallSign)
    gVessel.sInmarsat = NullTrim(gRs!sLloyd)
    gVessel.sLloyd = NullTrim(gRs!sLloyd)
    gVessel.nLoa = NullTrim(gRs!nLoa)
    gVessel.nLbp = NullTrim(gRs!nLbp)
    gVessel.nWidth = NullTrim(gRs!nWidth)
    gVessel.nDepth = NullTrim(gRs!nDepth)
    gVessel.nTopHgt = NullTrim(gRs!nTopHgt)
    gVessel.nAntHgt = NullTrim(gRs!nAntHgt)
    gVessel.iDMaxRows = NullTrim(gRs!IDMAXROW)
    gVessel.iDMaxTiers = NullTrim(gRs!IDMAXTIER)
    gVessel.iHMaxRows = NullTrim(gRs!IHMAXROW)
    gVessel.iHMaxTiers = NullTrim(gRs!IHMAXTIER)
    gVessel.iNoHatchs = NullTrim(gRs!INOHATCH)
    gVessel.iNoBays = NullTrim(gRs!iNoBay)
    gVessel.iBgNo = NullTrim(gRs!iBgNo)
    gVessel.nBgLength = NullTrim_Num(gRs!nBgLength)
    gVessel.sConfNm = NullTrim(gRs!INS_ID)
    
    If gVessel.iDMaxRows >= gVessel.iHMaxRows Then
      gVessel.iMaxRows = gVessel.iDMaxRows
    Else
      gVessel.iMaxRows = gVessel.iHMaxRows
    End If
    
    If gVessel.iDMaxTiers >= gVessel.iHMaxTiers Then
      gVessel.iMaxTiers = gVessel.iDMaxTiers
    Else
      gVessel.iMaxTiers = gVessel.iHMaxTiers
    End If
      
    
    ReDim Preserve gVessel.gHatch(1 To NullTrim_Num(gRs!INOHATCH))
    ReDim Preserve gVessel.gBay(1 To NullTrim_Num(gRs!iNoBay))
    
    gVessel.bSaveFlag = True
    
    gVesFlag = True
  End If
  gRs.Close
  
  If gVesFlag = False Then Exit Sub
  
  'HATCH
  SQL = "SELECT IHCH,"
  SQL = SQL & "        IDTIER,"
  SQL = SQL & "        IDFRROW,"
  SQL = SQL & "        IDTOROW,"
  SQL = SQL & "        IHTIER,"
  SQL = SQL & "        IHFRROW,"
  SQL = SQL & "        IHTOROW"
  SQL = SQL & "   FROM TB_VD_HATCH"
  SQL = SQL & "  WHERE SVCODE = '" & sVslCd & "'"
  SQL = SQL & "  ORDER BY IHCH"
  Set gRs = G_Host_Con.Execute(SQL)
  
  If Not (gRs.BOF And gRs.EOF) Then
    Do While Not gRs.EOF
      gVessel.gHatch(NullTrim_Num(gRs!iHch)).iHch = NullTrim_Num(gRs!iHch)
      gVessel.gHatch(NullTrim_Num(gRs!iHch)).sHchNo = Format(NullTrim(gRs!iHch), "00")
      gVessel.gHatch(NullTrim_Num(gRs!iHch)).iDTier = NullTrim_Num(gRs!iDTier)
      gVessel.gHatch(NullTrim_Num(gRs!iHch)).iDFrRow = NullTrim_Num(gRs!iDFrRow)
      gVessel.gHatch(NullTrim_Num(gRs!iHch)).iDToRow = NullTrim_Num(gRs!iDToRow)
      gVessel.gHatch(NullTrim_Num(gRs!iHch)).iHTier = NullTrim_Num(gRs!iHTier)
      gVessel.gHatch(NullTrim_Num(gRs!iHch)).iHFrRow = NullTrim_Num(gRs!iHFrRow)
      gVessel.gHatch(NullTrim_Num(gRs!iHch)).iHToRow = NullTrim_Num(gRs!iHToRow)
      
      gVessel.gHatch(NullTrim_Num(gRs!iHch)).bSaveFlag = True
      
      gRs.MoveNext
    Loop
  End If
  gRs.Close
  
  'Bay Info
  SQL = "SELECT IBAY, SBAYNO, SSIZE, IHCH, INOCVR"
  SQL = SQL & "   FROM TB_VD_BAY"
  SQL = SQL & "  WHERE SVCODE = '" & sVslCd & "'"
  SQL = SQL & "  ORDER BY IBAY"
  Set gRs = G_Host_Con.Execute(SQL)
  If Not (gRs.BOF And gRs.EOF) Then
    Do While Not gRs.EOF
      gVessel.gBay(NullTrim_Num(gRs!iBay)).iBay = NullTrim_Num(gRs!iBay)
      gVessel.gBay(NullTrim_Num(gRs!iBay)).sBayNo = NullTrim(gRs!sBayNo)
      gVessel.gBay(NullTrim_Num(gRs!iBay)).sSize = NullTrim(gRs!sSize)
      gVessel.gBay(NullTrim_Num(gRs!iBay)).iHchNo = NullTrim_Num(gRs!iHch)
      gVessel.gBay(NullTrim_Num(gRs!iBay)).iNoCvr = NullTrim_Num(gRs!iNoCvr)
      gVessel.gBay(NullTrim_Num(gRs!iBay)).iNoRows(0) = gVessel.iHMaxRows
      gVessel.gBay(NullTrim_Num(gRs!iBay)).iNoRows(1) = gVessel.iDMaxRows
      gVessel.gBay(NullTrim_Num(gRs!iBay)).iNoTiers(0) = gVessel.iHMaxTiers
      gVessel.gBay(NullTrim_Num(gRs!iBay)).iNoTiers(1) = gVessel.iDMaxTiers
      
      ReDim gVessel.gBay(NullTrim_Num(gRs!iBay)).gRow(0 To 1, 1 To gVessel.iMaxRows)
      ReDim gVessel.gBay(NullTrim_Num(gRs!iBay)).gTier(0 To 1, 1 To gVessel.iMaxTiers)
      ReDim gVessel.gBay(NullTrim_Num(gRs!iBay)).gCell(0 To 1, 1 To gVessel.iMaxRows, 1 To gVessel.iMaxTiers)
      ReDim gVessel.gBay(NullTrim_Num(gRs!iBay)).gCellOrg(0 To 1, 1 To gVessel.iMaxRows, 1 To gVessel.iMaxTiers)
      
      'Hatch Cover Info
      If gVessel.gBay(NullTrim_Num(gRs!iBay)).iNoCvr > 0 Then
        ReDim Preserve gVessel.gBay(NullTrim_Num(gRs!iBay)).gHchCvr(1 To gVessel.gBay(NullTrim_Num(gRs!iBay)).iNoCvr)
      End If

      'Hatch Cover Info
      SQL = "SELECT IBAY, ICVR, SHTYPE, NF, NT"
      SQL = SQL & "   FROM TB_VD_COVER"
      SQL = SQL & "  WHERE SVCODE = '" & sVslCd & "' AND IBAY = '" & NullTrim_Num(gRs!iBay) & "'"
      SQL = SQL & "  ORDER BY ICVR"
      Set tRs = G_Host_Con.Execute(SQL)
      If Not (tRs.BOF And tRs.EOF) Then
        Do While Not tRs.EOF
          gVessel.gBay(NullTrim_Num(gRs!iBay)).gHchCvr(NullTrim_Num(tRs!iCvr)).iCvr = NullTrim_Num(tRs!iCvr)
          gVessel.gBay(NullTrim_Num(gRs!iBay)).gHchCvr(NullTrim_Num(tRs!iCvr)).sHType = NullTrim_Num(tRs!sHType)
          gVessel.gBay(NullTrim_Num(gRs!iBay)).gHchCvr(NullTrim_Num(tRs!iCvr)).nF = NullTrim_Num(tRs!nF)
          gVessel.gBay(NullTrim_Num(gRs!iBay)).gHchCvr(NullTrim_Num(tRs!iCvr)).nT = NullTrim_Num(tRs!nT)
          gVessel.gBay(NullTrim_Num(gRs!iBay)).gHchCvr(NullTrim_Num(tRs!iCvr)).iBay = NullTrim_Num(tRs!iBay)

          tRs.MoveNext
        Loop
      End If
      tRs.Close
      
      'Row Info
      SQL = "SELECT IBAY, IHD, IROW, SNO, NSTKWGT"
      SQL = SQL & "   FROM TB_VD_ROW"
      SQL = SQL & "  WHERE SVCODE = '" & sVslCd & "' AND IBAY = " & NullTrim_Num(gRs!iBay)
      SQL = SQL & "  ORDER BY IHD, IROW"
      Set tRs = G_Host_Con.Execute(SQL)
      If Not (tRs.BOF And tRs.EOF) Then
        Do While Not tRs.EOF
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gRow(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow)).iHD = NullTrim_Num(tRs!iHD)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gRow(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow)).iRow = NullTrim_Num(tRs!iRow)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gRow(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow)).sNo = NullTrim(tRs!sNo)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gRow(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow)).nStkWgt = NullTrim_Num(tRs!nStkWgt)
'          gVessel.gBay(NullTrim_Num(tRs!iBay)).gRow(NullTrim_Num(tRs!iHd), NullTrim_Num(tRs!iRow)).nRowClr = NullTrim_Num(tRs!nRowClr)
          
          tRs.MoveNext
        Loop
      End If
      tRs.Close
      
      'Tier Info
      SQL = "SELECT IBAY, IHD, ITIER, SNO"
      SQL = SQL & "   FROM TB_VD_TIER"
      SQL = SQL & "  WHERE SVCODE = '" & sVslCd & "' AND IBAY = " & NullTrim_Num(gRs!iBay)
      SQL = SQL & "  ORDER BY IHD, ITIER"
      Set tRs = G_Host_Con.Execute(SQL)
      If Not (tRs.BOF And tRs.EOF) Then
        Do While Not tRs.EOF
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gTier(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iTier)).iHD = NullTrim_Num(tRs!iHD)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gTier(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iTier)).iTier = NullTrim_Num(tRs!iTier)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gTier(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iTier)).sNo = NullTrim(tRs!sNo)
          
          tRs.MoveNext
        Loop
      End If
      tRs.Close
      
      'Cell Info
      SQL = "SELECT IBAY, IHD, IROW, ITIER, SS, SGUIDE"
      SQL = SQL & "   FROM TB_VD_CELL"
      SQL = SQL & "  WHERE SVCODE = '" & sVslCd & "' AND IBAY = " & NullTrim_Num(gRs!iBay)
      Set tRs = G_Host_Con.Execute(SQL)
      If Not (tRs.BOF And tRs.EOF) Then
        Do While Not tRs.EOF
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCell(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).iHD = NullTrim_Num(tRs!iHD)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCell(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).iRow = NullTrim_Num(tRs!iRow)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCell(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).iTier = NullTrim_Num(tRs!iTier)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCell(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).sSS = NullTrim_Num(tRs!ss)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCell(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).sGuide = NullTrim_Num(tRs!sGuide)
          
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCellOrg(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).iHD = NullTrim_Num(tRs!iHD)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCellOrg(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).iRow = NullTrim_Num(tRs!iRow)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCellOrg(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).iTier = NullTrim_Num(tRs!iTier)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCellOrg(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).sSS = NullTrim(tRs!ss)
          gVessel.gBay(NullTrim_Num(tRs!iBay)).gCellOrg(NullTrim_Num(tRs!iHD), NullTrim_Num(tRs!iRow), NullTrim_Num(tRs!iTier)).sGuide = NullTrim(tRs!sGuide)
          
          tRs.MoveNext
        Loop
      End If
      tRs.Close
      
      gVessel.gBay(NullTrim_Num(gRs!iBay)).bSaveFlag = True
      
      gRs.MoveNext
    Loop
  End If
  
  gRs.Close: Set gRs = Nothing
  
  Screen.MousePointer = 0
End Sub

Public Function SaveVslStructure()
  Dim i%, j%, k%, l%
  
  Screen.MousePointer = 11
  
  'VESSEL INFO
  If gVessel.bSaveFlag Then
    
    SQL = "DELETE FROM TB_VD_VESSEL WHERE SVCODE = '" & gVessel.sVCode & "'"
    G_Host_Con.Execute (SQL)
    
    SQL = "INSERT INTO TB_VD_VESSEL"
    SQL = SQL & "   (SVCODE,"
    SQL = SQL & "    SVNAME,"
    SQL = SQL & "    SCALLSIGN,"
    SQL = SQL & "    SINMARSAT,"
    SQL = SQL & "    SLLOYD,"
    SQL = SQL & "    NLOA,"
    SQL = SQL & "    NLBP,"
    SQL = SQL & "    NWIDTH,"
    SQL = SQL & "    NDEPTH,"
    SQL = SQL & "    NTOPHGT,"
    SQL = SQL & "    NANTHGT,"
    SQL = SQL & "    INOBAY,"
    SQL = SQL & "    INOHATCH,"
    SQL = SQL & "    IDMAXROW,"
    SQL = SQL & "    IHMAXROW,"
    SQL = SQL & "    IDMAXTIER,"
    SQL = SQL & "    IHMAXTIER,"
    SQL = SQL & "    IBGNO,"
'    SQL = SQL & "    NBGLENGTH,"
    SQL = SQL & "    INS_ID,"
    SQL = SQL & "    INS_DT) VALUES"
    SQL = SQL & "   ('" & gVessel.sVCode & "',"
    SQL = SQL & "    '" & gVessel.sVName & "',"
    SQL = SQL & "    '" & gVessel.sCallSign & "',"
    SQL = SQL & "    '" & gVessel.sInmarsat & "',"
    SQL = SQL & "    '" & "" & "',"
    SQL = SQL & "    '" & gVessel.nLoa & "',"
    SQL = SQL & "    '" & gVessel.nLbp & "',"
    SQL = SQL & "    '" & gVessel.nWidth & "',"
    SQL = SQL & "    '" & gVessel.nDepth & "',"
    SQL = SQL & "    '" & gVessel.nTopHgt & "',"
    SQL = SQL & "    '" & gVessel.nAntHgt & "',"
    SQL = SQL & "    '" & gVessel.iNoBays & "',"
    SQL = SQL & "    '" & gVessel.iNoHatchs & "',"
    SQL = SQL & "    '" & gVessel.iDMaxRows & "',"
    SQL = SQL & "    '" & gVessel.iHMaxRows & "',"
    SQL = SQL & "    '" & gVessel.iDMaxTiers & "',"
    SQL = SQL & "    '" & gVessel.iHMaxTiers & "',"
    SQL = SQL & "    '" & gVessel.iBgNo & "',"
'    SQL = SQL & "    '" & gVessel.nBgLength & "',"
    SQL = SQL & "    '" & gVessel.sConfNm & "',"
    SQL = SQL & "    CURRENT_TIMESTAMP)"
    G_Host_Con.Execute (SQL)
  End If
  
  'HATCH INFO
  SQL = "DELETE FROM TB_VD_HATCH WHERE SVCODE = '" & gVessel.sVCode & "'"
  G_Host_Con.Execute (SQL)
  For i = 1 To gVessel.iNoHatchs
    SQL = "INSERT INTO TB_VD_HATCH"
    SQL = SQL & "   (SVCODE,"
    SQL = SQL & "    IHCH,"
'    SQL = SQL & "    INOCVR,"
    SQL = SQL & "    IDTIER,"
    SQL = SQL & "    IDFRROW,"
    SQL = SQL & "    IDTOROW,"
    SQL = SQL & "    IHTIER,"
    SQL = SQL & "    IHFRROW,"
    SQL = SQL & "    IHTOROW) VALUES"
    SQL = SQL & "   ('" & gVessel.sVCode & "',"
    SQL = SQL & "    '" & i & "',"
'    SQL = SQL & "    '" & gVessel.gHatch(i).iNoCvr & "',"
    SQL = SQL & "    '" & gVessel.gHatch(i).iDTier & "',"
    SQL = SQL & "    '" & gVessel.gHatch(i).iDFrRow & "',"
    SQL = SQL & "    '" & gVessel.gHatch(i).iDToRow & "',"
    SQL = SQL & "    '" & gVessel.gHatch(i).iHTier & "',"
    SQL = SQL & "    '" & gVessel.gHatch(i).iHFrRow & "',"
    SQL = SQL & "    '" & gVessel.gHatch(i).iHToRow & "')"
    G_Host_Con.Execute (SQL)
  Next i
  
  'BAY INFO
  SQL = "DELETE FROM TB_VD_BAY WHERE SVCODE = '" & gVessel.sVCode & "'"
  G_Host_Con.Execute (SQL)
  For i = 1 To gVessel.iNoBays
    SQL = "INSERT INTO TB_VD_BAY"
    SQL = SQL & "   (SVCODE,"
    SQL = SQL & "    IBAY,"
    SQL = SQL & "    SBAYNO,"
    SQL = SQL & "    SSIZE,"
    SQL = SQL & "    IHCH,"
    SQL = SQL & "    INOCVR) VALUES"
    SQL = SQL & "   ('" & gVessel.sVCode & "',"
    SQL = SQL & "    '" & i & "',"
    SQL = SQL & "    '" & gVessel.gBay(i).sBayNo & "',"
    SQL = SQL & "    '" & gVessel.gBay(i).sSize & "',"
    SQL = SQL & "    '" & gVessel.gBay(i).iHchNo & "',"
    SQL = SQL & "    '" & gVessel.gBay(i).iNoCvr & "')"
    G_Host_Con.Execute (SQL)
    
    'HATCH COVER INFO
    SQL = "DELETE FROM TB_VD_COVER WHERE SVCODE = '" & gVessel.sVCode & "' AND IBAY = " & i
    G_Host_Con.Execute (SQL)
    
    For j = 1 To gVessel.gBay(i).iNoCvr
      SQL = "INSERT INTO TB_VD_COVER"
      SQL = SQL & "   (SVCODE,"
      SQL = SQL & "    IBAY,"
      SQL = SQL & "    ICVR,"
      SQL = SQL & "    SHTYPE,"
      SQL = SQL & "    NF,"
      SQL = SQL & "    NT) VALUES"
      SQL = SQL & "   ('" & gVessel.sVCode & "',"
      SQL = SQL & "    '" & i & "',"
      SQL = SQL & "    '" & j & "',"
      SQL = SQL & "    '" & gVessel.gBay(i).gHchCvr(j).sHType & "',"
      SQL = SQL & "    '" & gVessel.gBay(i).gHchCvr(j).nF & "',"
      SQL = SQL & "    '" & gVessel.gBay(i).gHchCvr(j).nT & "')"
      G_Host_Con.Execute (SQL)
    Next j
    
    'ROW INFO
    SQL = "DELETE FROM TB_VD_ROW WHERE SVCODE = '" & gVessel.sVCode & "' AND IBAY = " & gVessel.gBay(i).iBay
    G_Host_Con.Execute (SQL)
    
    For j = 0 To 1
      For k = 1 To gVessel.gBay(i).iNoRows(j)
        'If NullTrim(gVessel.gBay(i).gRow(j, k).sNo) <> "" Then
          SQL = "INSERT INTO TB_VD_ROW"
          SQL = SQL & "   (SVCODE,"
          SQL = SQL & "    IBAY,"
          SQL = SQL & "    IHD,"
          SQL = SQL & "    IROW,"
          SQL = SQL & "    SNO,"
          SQL = SQL & "    NSTKWGT) VALUES"
          SQL = SQL & "   ('" & gVessel.sVCode & "',"
          SQL = SQL & "    '" & i & "',"
          SQL = SQL & "    '" & j & "',"
          SQL = SQL & "    '" & k & "',"
          SQL = SQL & "    '" & NullTrim(gVessel.gBay(i).gRow(j, k).sNo) & "',"
          SQL = SQL & "    '" & gVessel.gBay(i).gRow(j, k).nStkWgt & "')"
          G_Host_Con.Execute (SQL)
        'End If
      Next k
    Next j
    
    'TIER INFO
    SQL = "DELETE FROM TB_VD_TIER WHERE SVCODE = '" & gVessel.sVCode & "' AND IBAY = " & gVessel.gBay(i).iBay
    G_Host_Con.Execute (SQL)
    
    For j = 0 To 1
      For k = 1 To gVessel.gBay(i).iNoTiers(j)
        'If NullTrim(gVessel.gBay(i).gTier(j, k).sNo) <> "" Then
          SQL = "INSERT INTO TB_VD_TIER"
          SQL = SQL & "   (SVCODE,"
          SQL = SQL & "    IBAY,"
          SQL = SQL & "    IHD,"
          SQL = SQL & "    ITIER,"
          SQL = SQL & "    SNO) VALUES"
          SQL = SQL & "   ('" & gVessel.sVCode & "',"
          SQL = SQL & "    '" & i & "',"
          SQL = SQL & "    '" & j & "',"
          SQL = SQL & "    '" & k & "',"
          SQL = SQL & "    '" & NullTrim(gVessel.gBay(i).gTier(j, k).sNo) & "')"
          G_Host_Con.Execute (SQL)
        'End If
      Next k
    Next j
    
    'CELL INFO
    SQL = "DELETE FROM TB_VD_CELL WHERE SVCODE = '" & gVessel.sVCode & "' AND IBAY = " & gVessel.gBay(i).iBay
    G_Host_Con.Execute (SQL)
    
    For j = 0 To 1
      For k = 1 To gVessel.gBay(i).iNoRows(j)
        For l = 1 To gVessel.gBay(i).iNoTiers(j)
          'If gVessel.gBay(i).gCell(j, k, l).sSS = "Y" Then
            SQL = "INSERT INTO TB_VD_CELL"
            SQL = SQL & "   (SVCODE,"
            SQL = SQL & "    IBAY,"
            SQL = SQL & "    IHD,"
            SQL = SQL & "    IROW,"
            SQL = SQL & "    ITIER,"
            SQL = SQL & "    SS,"
            SQL = SQL & "    SGUIDE) VALUES"
            SQL = SQL & "   ('" & gVessel.sVCode & "',"
            SQL = SQL & "    '" & i & "',"
            SQL = SQL & "    '" & j & "',"
            SQL = SQL & "    '" & k & "',"
            SQL = SQL & "    '" & l & "',"
            SQL = SQL & "    '" & NullTrim(gVessel.gBay(i).gCell(j, k, l).sSS) & "',"
            SQL = SQL & "    '" & NullTrim(gVessel.gBay(i).gCell(j, k, l).sGuide) & "')"
            G_Host_Con.Execute (SQL)
          'End If
        Next l
      Next k
    Next j
  Next i
  
  MsgBox "Save Vessel Defined Info. Successfully!"
  Screen.MousePointer = 0
End Function

Public Function ChkVslCd(sVslCd$) As Boolean
  ChkVslCd = False
  
  SQL = "SELECT * FROM TB_VD_VESSEL WHERE SVCODE = '" & sVslCd & "'"
  Set gRs = G_Host_Con.Execute(SQL)
  
  If Not (gRs.BOF And gRs.EOF) Then
    ChkVslCd = True
  End If
  gRs.Close: Set gRs = Nothing
End Function

Public Sub DelVesselInfo(sVslCd$)
  SQL = "DELETE FROM TB_VD_VESSEL WHERE SVCODE = '" & sVslCd & "'"
  G_Host_Con.Execute (SQL)
  
  SQL = "DELETE FROM TB_VD_HATCH WHERE SVCODE = '" & sVslCd & "'"
  G_Host_Con.Execute (SQL)
  
  SQL = "DELETE FROM TB_VD_COVER WHERE SVCODE = '" & sVslCd & "'"
  G_Host_Con.Execute (SQL)
  
  SQL = "DELETE FROM TB_VD_BAY WHERE SVCODE = '" & sVslCd & "'"
  G_Host_Con.Execute (SQL)
  
  SQL = "DELETE FROM TB_VD_ROW WHERE SVCODE = '" & sVslCd & "'"
  G_Host_Con.Execute (SQL)
  
  SQL = "DELETE FROM TB_VD_TIER WHERE SVCODE = '" & sVslCd & "'"
  G_Host_Con.Execute (SQL)
  
  SQL = "DELETE FROM TB_VD_CELL WHERE SVCODE = '" & sVslCd & "'"
  G_Host_Con.Execute (SQL)
End Sub

Public Sub SetSvrConfig()
  Dim fso As New FileSystemObject
  Dim strPath$, f
  
  strPath = App.Path & "\" & cSvrConfigFile
  If fso.FileExists(strPath) = True Then
    f = FreeFile
    Open App.Path & "\" & cSvrConfigFile For Input As f
    Input #f, gSvrCfg.SvrIp
    Input #f, gSvrCfg.SvrId
    Input #f, gSvrCfg.SvrPw
    Input #f, gSvrCfg.SvrDb
    Close #f
  End If
End Sub
