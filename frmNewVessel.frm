VERSION 5.00
Begin VB.Form frmNewVessel 
   BorderStyle     =   1  '¥‹¿œ ∞Ì¡§
   Caption         =   "Create a New Vessel - General Information"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmNewVessel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7920
   StartUpPosition =   2  '»≠∏È ∞°øÓµ•
   Begin VB.Frame fraBasic 
      BorderStyle     =   0  'æ¯¿Ω
      Height          =   855
      Left            =   0
      TabIndex        =   37
      Top             =   3600
      Width           =   7815
      Begin VB.CommandButton cmdBack 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5040
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   6360
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line lSep 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   7560
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraGeneral 
      BorderStyle     =   0  'æ¯¿Ω
      Height          =   3495
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7815
      Begin VB.Frame Frame1 
         Caption         =   "General Information"
         Height          =   3375
         Left            =   2040
         TabIndex        =   24
         Top             =   120
         Width           =   5775
         Begin VB.TextBox txtVslCd 
            Height          =   270
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   0
            Top             =   300
            Width           =   1575
         End
         Begin VB.TextBox txtLoa 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   5
            Text            =   "0"
            Top             =   2220
            Width           =   975
         End
         Begin VB.TextBox txtWidth 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "0"
            Top             =   2580
            Width           =   975
         End
         Begin VB.TextBox txtTopHgt 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   9
            Text            =   "0"
            Top             =   2940
            Width           =   975
         End
         Begin VB.TextBox txtLbp 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   6
            Text            =   "0"
            Top             =   2220
            Width           =   975
         End
         Begin VB.TextBox txtDepth 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   4560
            MaxLength       =   4
            TabIndex        =   8
            Text            =   "0"
            Top             =   2580
            Width           =   975
         End
         Begin VB.TextBox txtAntHgt 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   4560
            MaxLength       =   4
            TabIndex        =   10
            Text            =   "0"
            Top             =   2940
            Width           =   975
         End
         Begin VB.TextBox txtVslName 
            Height          =   270
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   1
            Top             =   680
            Width           =   3735
         End
         Begin VB.TextBox txtCallSign 
            Height          =   270
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   2
            Top             =   1020
            Width           =   1575
         End
         Begin VB.TextBox txtInmarsal 
            Height          =   270
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   3
            Top             =   1410
            Width           =   1575
         End
         Begin VB.TextBox txtLloyd 
            Height          =   270
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1770
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Vessel Code :"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Vessel Name :"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   220
            TabIndex        =   34
            Top             =   720
            Width           =   1470
         End
         Begin VB.Label Label3 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Call Sign :"
            Height          =   180
            Left            =   360
            TabIndex        =   33
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Inmarsat No. :"
            Height          =   180
            Left            =   480
            TabIndex        =   32
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Lloyd's Code :"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Top Tier Height :"
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   240
            X2              =   5520
            Y1              =   2130
            Y2              =   2130
         End
         Begin VB.Label Label7 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "L.O.A :"
            Height          =   180
            Left            =   720
            TabIndex        =   29
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label8 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Width :"
            Height          =   180
            Left            =   720
            TabIndex        =   28
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label9 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Antenna Height :"
            Height          =   180
            Left            =   3000
            TabIndex        =   27
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label Label10 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "L.B.P :"
            Height          =   180
            Left            =   3600
            TabIndex        =   26
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label11 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Depth :"
            Height          =   180
            Left            =   3600
            TabIndex        =   25
            Top             =   2640
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000080&
         BorderStyle     =   0  'æ¯¿Ω
         Height          =   200
         Left            =   1560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   36
         Top             =   2280
         Width           =   200
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'æ¯¿Ω
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmNewVessel.frx":038A
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   1665
         Left            =   240
         Picture         =   "frmNewVessel.frx":03DB
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Frame fraStructure 
      BorderStyle     =   0  'æ¯¿Ω
      Height          =   3495
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   7815
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000080&
         BorderStyle     =   0  'æ¯¿Ω
         Height          =   200
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   39
         Top             =   2280
         Width           =   200
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'æ¯¿Ω
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   42
         Text            =   "frmNewVessel.frx":14FA
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Structure Information"
         Height          =   3375
         Left            =   1920
         TabIndex        =   40
         Top             =   120
         Width           =   5775
         Begin VB.TextBox txtBayCnt 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   4800
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "0"
            Top             =   960
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtBgNo 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   4800
            MaxLength       =   2
            TabIndex        =   17
            Text            =   "0"
            Top             =   1860
            Width           =   735
         End
         Begin VB.TextBox txtHoldTiers 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   4800
            MaxLength       =   2
            TabIndex        =   14
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtDeckTiers 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   4800
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtHatchCnt 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   15
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtHoldRows 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtDeckRows 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   11
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtBgLength 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Height          =   270
            Left            =   4800
            MaxLength       =   5
            TabIndex        =   18
            Text            =   "0"
            Top             =   2220
            Width           =   735
         End
         Begin VB.Label Label16 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Bay Count :"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   3240
            TabIndex        =   50
            Top             =   1020
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label18 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            AutoSize        =   -1  'True
            Caption         =   "Bridge"
            Height          =   180
            Left            =   240
            TabIndex        =   49
            Top             =   1560
            Width           =   540
         End
         Begin VB.Label Label17 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Index of the Hatch located before Bridge :"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   120
            TabIndex        =   48
            Top             =   1920
            Width           =   4575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Max Hold Tiers :"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   3000
            TabIndex        =   47
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label Label15 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Max Deck Tiers :"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   2880
            TabIndex        =   46
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label Label14 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Max Hold Rows :"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   120
            TabIndex        =   45
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label Label13 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Max Deck Rows :"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   75
            TabIndex        =   44
            Top             =   300
            Width           =   1740
         End
         Begin VB.Label Label12 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "Hatch Count :"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   120
            TabIndex        =   43
            Top             =   1020
            Width           =   1695
         End
         Begin VB.Label Label21 
            Alignment       =   1  'ø¿∏•¬  ∏¬√„
            Caption         =   "A point of distance from stern for Bitt position (m) :"
            Height          =   180
            Left            =   360
            TabIndex        =   41
            Top             =   2280
            Width           =   4335
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   240
            X2              =   5520
            Y1              =   1440
            Y2              =   1440
         End
      End
      Begin VB.Image Image2 
         Height          =   1725
         Left            =   120
         Picture         =   "frmNewVessel.frx":154B
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmNewVessel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
  fraGeneral.Visible = True
  fraStructure.Visible = False
  cmdBack.Enabled = False
  cmdNext.Enabled = True
  Me.Caption = "Create a New Vessel - General Information"
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdNext_Click()
  Dim i%, j%, idx%
  Dim iVal%, iVal2%
  
  If fraStructure.Visible Then
    Erase gVessel.gHatch
    Erase gVessel.gBay
    
    With gVessel
      'Vessel Info
      .sVCode = txtVslCd.Text
      .sVName = txtVslName.Text
      .sCallSign = txtCallSign.Text
      .sInmarsat = txtInmarsal.Text
      .sLloyd = txtLloyd.Text
      .nLoa = Val(txtLoa.Text)
      .nLbp = Val(txtLbp.Text)
      .nWidth = Val(txtWidth.Text)
      .nDepth = Val(txtDepth.Text)
      .nTopHgt = Val(txtTopHgt.Text)
      .nAntHgt = Val(txtAntHgt.Text)
      .iNoBays = Val(txtHatchCnt.Text) * 3
      .iNoHatchs = Val(txtHatchCnt.Text)
      .iDMaxRows = Val(txtDeckRows.Text)
      .iHMaxRows = Val(txtHoldRows.Text)
      .iDMaxTiers = Val(txtDeckTiers.Text)
      .iHMaxTiers = Val(txtHoldTiers.Text)
      .iBgNo = Val(txtBgNo.Text)
      .nBgLength = Val(txtBgLength.Text)
      .bSaveFlag = True
      
      If .iHMaxRows >= .iDMaxRows Then
        .iMaxRows = .iHMaxRows
      Else
        .iMaxRows = .iDMaxRows
      End If
      
      If .iHMaxTiers >= .iDMaxTiers Then
        .iMaxTiers = .iHMaxTiers
      Else
        .iMaxTiers = .iDMaxTiers
      End If
      
      'Hatch Info
      ReDim Preserve .gHatch(1 To .iNoHatchs)
      For i = 1 To .iNoHatchs
        .gHatch(i).iHch = i
        .gHatch(i).sHchNo = Format(i, "00")
        .gHatch(i).iDTier = .iDMaxTiers
        .gHatch(i).iHTier = 1
        .gHatch(i).iDFrRow = 1
        .gHatch(i).iDToRow = .iDMaxRows
        .gHatch(i).iHFrRow = 1
        .gHatch(i).iHToRow = .iHMaxRows
        
        .gHatch(i).bSaveFlag = True
      Next i
      
      'Bay Info
      ReDim Preserve .gBay(1 To .iNoBays)
      For i = 1 To .iNoBays
        .gBay(i).iBay = i
        idx = idx + 1
        If idx Mod 4 = 0 Then idx = idx + 1
        .gBay(i).sBayNo = Format(idx, "00")
        .gBay(i).iHchNo = Int((i - 1) / 3) + 1
        
        If (i + 1) Mod 3 = 0 Then
          .gBay(i).sSize = "4"
        Else
          .gBay(i).sSize = "2"
        End If
        
        .gBay(i).bSaveFlag = True
        
        'Row Info
        .gBay(i).iNoRows(0) = .iHMaxRows
        .gBay(i).iNoRows(1) = .iDMaxRows
        ReDim .gBay(i).gRow(0 To 1, 1 To .iMaxRows)
        For j = 1 To .iHMaxRows
          .gBay(i).gRow(0, j).iHD = 0
          .gBay(i).gRow(0, j).iRow = j
          
          iVal = Int(.iHMaxRows / 2)
          If .iHMaxRows Mod 2 = 0 Then
            If j <= iVal Then
              iVal2 = .iHMaxRows - (j - 1) * 2
              .gBay(i).gRow(0, j).sNo = Format(iVal2, "00")
            Else
              iVal2 = (j - iVal) * 2 - 1
              .gBay(i).gRow(0, j).sNo = Format(iVal2, "00")
            End If
            
          Else
            If j <= iVal Then
              iVal2 = (.iHMaxRows - 1) - (j - 1) * 2
              .gBay(i).gRow(0, j).sNo = Format(iVal2, "00")
            ElseIf j = iVal + 1 Then
              .gBay(i).gRow(0, j).sNo = "00"
            Else
              iVal2 = (j - iVal - 1) * 2 - 1
              .gBay(i).gRow(0, j).sNo = Format(iVal2, "00")
            End If
            
          End If
        Next j
        
        For j = 1 To .iDMaxRows
          .gBay(i).gRow(1, j).iHD = 1
          .gBay(i).gRow(1, j).iRow = j
          
          iVal = Int(.iDMaxRows / 2)
          If .iDMaxRows Mod 2 = 0 Then
            If j <= iVal Then
              iVal2 = .iDMaxRows - (j - 1) * 2
              .gBay(i).gRow(1, j).sNo = Format(iVal2, "00")
            Else
              iVal2 = (j - iVal) * 2 - 1
              .gBay(i).gRow(1, j).sNo = Format(iVal2, "00")
            End If
            
          Else
            If j <= iVal Then
              iVal2 = (.iDMaxRows - 1) - (j - 1) * 2
              .gBay(i).gRow(1, j).sNo = Format(iVal2, "00")
            ElseIf j = iVal + 1 Then
              .gBay(i).gRow(1, j).sNo = "00"
            Else
              iVal2 = (j - iVal - 1) * 2 - 1
              .gBay(i).gRow(1, j).sNo = Format(iVal2, "00")
            End If
            
          End If
        Next j
        
          
        'Tier Info
        .gBay(i).iNoTiers(0) = gVessel.iHMaxTiers
        .gBay(i).iNoTiers(1) = gVessel.iDMaxTiers
        ReDim .gBay(i).gTier(0 To 1, 1 To .iMaxTiers)
        For j = 1 To .iHMaxTiers
          .gBay(i).gTier(0, j).iHD = 0
          .gBay(i).gTier(0, j).iTier = j
          iVal = j * 2
          .gBay(i).gTier(0, j).sNo = Format(iVal, "00")
        Next j
        
        For j = 1 To .iDMaxTiers
          .gBay(i).gTier(1, j).iHD = 1
          .gBay(i).gTier(1, j).iTier = j
          iVal = j * 2 + 80
          .gBay(i).gTier(1, j).sNo = Format(iVal, "00")
        Next j
        
        'Cell Info
        ReDim .gBay(i).gCell(0 To 1, 1 To .iMaxRows, 1 To .iMaxTiers)
        ReDim .gBay(i).gCellOrg(0 To 1, 1 To .iMaxRows, 1 To .iMaxTiers)
        
      Next i
    End With

    gVesFlag = True
    Call SetBasic
    
    Unload Me
  Else
'    If ChkVslCd(txtVslCd.Text) Then
'      MsgBox "Vessel Code is Exist!"
'      txtVslCd.SetFocus
'      Exit Sub
'    End If
    
    fraGeneral.Visible = False
    fraStructure.Visible = True
    cmdBack.Enabled = True
    Me.Caption = "Create a New Vessel - Structure Information"
    Call chkSndPhase
    txtDeckRows.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Me.Width = 7920
  Me.Height = 4800
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mdiVesConfig.Enabled = True
End Sub

Private Sub txtBgLength_GotFocus()
  txtBgLength.SelLength = Len(txtBgLength.Text)
End Sub

Private Sub txtBgNo_GotFocus()
  txtBgNo.SelLength = Len(txtBgNo.Text)
End Sub

Private Sub txtDeckRows_GotFocus()
  txtDeckRows.SelLength = Len(txtDeckRows.Text)
End Sub

Private Sub txtDeckRows_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkSndPhase
End Sub

Private Sub txtDeckTiers_GotFocus()
  txtDeckTiers.SelLength = Len(txtDeckTiers.Text)
End Sub

Private Sub txtDeckTiers_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkSndPhase
End Sub

Private Sub txtHatchCnt_GotFocus()
  txtHatchCnt.SelLength = Len(txtHatchCnt.Text)
End Sub

Private Sub txtHatchCnt_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkSndPhase
End Sub

Private Sub txtHoldRows_GotFocus()
  txtHoldRows.SelLength = Len(txtHoldRows.Text)
End Sub

Private Sub txtHoldRows_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkSndPhase
End Sub

Private Sub txtHoldTiers_GotFocus()
  txtHoldTiers.SelLength = Len(txtHoldTiers.Text)
End Sub

Private Sub txtHoldTiers_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkSndPhase
End Sub

Private Sub txtBgNo_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkSndPhase
End Sub

Private Sub txtVslCd_KeyPress(KeyAscii As Integer)
  If KeyAscii > 96 And KeyAscii < 124 Then             '¥ÎπÆ¿⁄∑Œ ∫Ø»Ø
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub txtVslCd_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkFstPhase
End Sub

Private Sub txtVslCd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call chkFstPhase
End Sub

Private Sub txtVslName_KeyPress(KeyAscii As Integer)
  If KeyAscii > 96 And KeyAscii < 124 Then             '¥ÎπÆ¿⁄∑Œ ∫Ø»Ø
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub txtVslName_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkFstPhase
End Sub

Private Sub chkFstPhase()
  If Trim(txtVslCd.Text) <> "" And Trim(txtVslName.Text) <> "" Then
    cmdNext.Enabled = True
  Else
    cmdNext.Enabled = False
  End If
End Sub

Private Sub chkSndPhase()
  If Val(txtDeckRows.Text) <> 0 And Val(txtDeckTiers.Text) <> 0 And Val(txtHoldRows.Text) <> 0 And Val(txtHoldTiers.Text) <> 0 And Val(txtHatchCnt.Text) <> 0 And Val(txtBgNo.Text) <> 0 Then
    cmdNext.Enabled = True
  Else
    cmdNext.Enabled = False
  End If
End Sub

Private Sub txtVslName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call chkFstPhase
End Sub
