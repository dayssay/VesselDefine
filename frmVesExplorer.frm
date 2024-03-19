VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmVesExplorer 
   Caption         =   "Vessel Explorer"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13830
   FillColor       =   &H00FF0000&
   Icon            =   "frmVesExplorer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   541
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   922
   WindowState     =   2  '최대화
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   3120
      ScaleHeight     =   461
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   669
      TabIndex        =   20
      Top             =   360
      Width           =   10095
      Begin VB.Frame fraBasic 
         BorderStyle     =   0  '없음
         Height          =   6735
         Left            =   840
         TabIndex        =   58
         Top             =   6000
         Visible         =   0   'False
         Width           =   4095
         Begin MSComctlLib.ImageList imgBayList 
            Left            =   1200
            Top             =   5880
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmVesExplorer.frx":058A
                  Key             =   "imgInsert"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmVesExplorer.frx":0B24
                  Key             =   "imgRemove"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar tbrBay 
            Height          =   330
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Appearance      =   1
            Style           =   1
            ImageList       =   "imgBayList"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "imgInsert"
                  Object.ToolTipText     =   "Insert Bay"
                  ImageKey        =   "imgInsert"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "imgRemove"
                  Object.ToolTipText     =   "Remove Bay"
                  ImageKey        =   "imgRemove"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
         End
         Begin FPSpread.vaSpread sprList 
            Height          =   12240
            Left            =   0
            TabIndex        =   59
            Top             =   340
            Width           =   3660
            _Version        =   196608
            _ExtentX        =   6456
            _ExtentY        =   21590
            _StockProps     =   64
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   100
            ScrollBars      =   2
            SpreadDesigner  =   "frmVesExplorer.frx":10BE
         End
      End
      Begin VB.Frame fraTable 
         BorderStyle     =   0  '없음
         Height          =   6375
         Left            =   3960
         TabIndex        =   32
         Top             =   5520
         Visible         =   0   'False
         Width           =   6375
         Begin VB.Frame Frame2 
            Caption         =   "Structure Information"
            Height          =   3855
            Left            =   120
            TabIndex        =   41
            Top             =   2400
            Width           =   6135
            Begin VB.TextBox txtBgLength 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   17
               Text            =   "0"
               Top             =   3360
               Width           =   735
            End
            Begin VB.TextBox txtBgNo 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   5040
               MaxLength       =   2
               TabIndex        =   16
               Text            =   "0"
               Top             =   3000
               Width           =   735
            End
            Begin VB.TextBox txtAntHgt 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   4800
               MaxLength       =   4
               TabIndex        =   15
               Text            =   "0"
               Top             =   2400
               Width           =   975
            End
            Begin VB.TextBox txtDepth 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   4800
               MaxLength       =   4
               TabIndex        =   13
               Text            =   "0"
               Top             =   2040
               Width           =   975
            End
            Begin VB.TextBox txtLbp 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   4800
               MaxLength       =   5
               TabIndex        =   11
               Text            =   "0"
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox txtTopHgt 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   1680
               MaxLength       =   4
               TabIndex        =   14
               Text            =   "0"
               Top             =   2400
               Width           =   975
            End
            Begin VB.TextBox txtWidth 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   1680
               MaxLength       =   4
               TabIndex        =   12
               Text            =   "0"
               Top             =   2040
               Width           =   975
            End
            Begin VB.TextBox txtLoa 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   1680
               MaxLength       =   5
               TabIndex        =   10
               Text            =   "0"
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox txtLloyd 
               Height          =   270
               Left            =   1680
               MaxLength       =   10
               TabIndex        =   9
               Top             =   1110
               Width           =   1335
            End
            Begin VB.TextBox txtInmarsal 
               Height          =   270
               Left            =   4560
               MaxLength       =   10
               TabIndex        =   8
               Top             =   750
               Width           =   1335
            End
            Begin VB.TextBox txtCallSign 
               Height          =   270
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   7
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtVslName 
               Height          =   270
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   6
               Top             =   240
               Width           =   4215
            End
            Begin VB.Label Label21 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "A point of distance from stern for Bitt position (m) :"
               Height          =   180
               Left            =   600
               TabIndex        =   53
               Top             =   3420
               Width           =   4335
            End
            Begin VB.Label Label17 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Index of the Hatch located before Bridge :"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   360
               TabIndex        =   52
               Top             =   3060
               Width           =   4575
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00808080&
               X1              =   120
               X2              =   5880
               Y1              =   2880
               Y2              =   2880
            End
            Begin VB.Label Label11 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Depth :"
               Height          =   180
               Left            =   3840
               TabIndex        =   51
               Top             =   2100
               Width           =   855
            End
            Begin VB.Label Label10 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "L.B.P :"
               Height          =   180
               Left            =   3840
               TabIndex        =   50
               Top             =   1740
               Width           =   855
            End
            Begin VB.Label Label9 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Antenna Height :"
               Height          =   180
               Left            =   3240
               TabIndex        =   49
               Top             =   2460
               Width           =   1455
            End
            Begin VB.Label Label8 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Width :"
               Height          =   180
               Left            =   720
               TabIndex        =   48
               Top             =   2100
               Width           =   855
            End
            Begin VB.Label Label7 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "L.O.A :"
               Height          =   180
               Left            =   720
               TabIndex        =   47
               Top             =   1740
               Width           =   855
            End
            Begin VB.Label Label6 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Top Tier Height :"
               Height          =   180
               Left            =   120
               TabIndex        =   46
               Top             =   2460
               Width           =   1455
            End
            Begin VB.Label Label5 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Lloyd's Code :"
               Height          =   180
               Left            =   0
               TabIndex        =   45
               Top             =   1140
               Width           =   1575
            End
            Begin VB.Label Label4 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Inmarsat No. :"
               Height          =   180
               Left            =   3240
               TabIndex        =   44
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label3 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Call Sign :"
               Height          =   180
               Left            =   240
               TabIndex        =   43
               Top             =   780
               Width           =   1335
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Vessel Name :"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   120
               TabIndex        =   42
               Top             =   285
               Width           =   1455
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               X1              =   120
               X2              =   5880
               Y1              =   1560
               Y2              =   1560
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Critical Information"
            Height          =   2175
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   6135
            Begin VB.CommandButton cmdChage 
               Caption         =   "&Change"
               Height          =   375
               Left            =   4680
               TabIndex        =   40
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtDeckRows 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   1
               Text            =   "0"
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox txtHoldRows 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   3
               Text            =   "0"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox txtHatchCnt 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   5
               Text            =   "0"
               Top             =   1680
               Width           =   735
            End
            Begin VB.TextBox txtDeckTiers 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   5040
               MaxLength       =   2
               TabIndex        =   2
               Text            =   "0"
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox txtHoldTiers 
               Alignment       =   1  '오른쪽 맞춤
               Height          =   270
               Left            =   5040
               MaxLength       =   2
               TabIndex        =   4
               Text            =   "0"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox txtVslCd 
               Height          =   270
               Left            =   1920
               MaxLength       =   4
               TabIndex        =   0
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Hatch Count :"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   120
               TabIndex        =   39
               Top             =   1740
               Width           =   1695
            End
            Begin VB.Label Label13 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Max Deck Rows :"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   240
               TabIndex        =   38
               Top             =   1020
               Width           =   1575
            End
            Begin VB.Label Label14 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Max Hold Rows :"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   240
               TabIndex        =   37
               Top             =   1380
               Width           =   1575
            End
            Begin VB.Label Label15 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Max Deck Tiers :"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   3360
               TabIndex        =   36
               Top             =   1020
               Width           =   1575
            End
            Begin VB.Label lbl 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Max Hold Tiers :"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   3480
               TabIndex        =   35
               Top             =   1380
               Width           =   1455
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00808080&
               X1              =   120
               X2              =   5880
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "Vessel Code :"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   360
               TabIndex        =   34
               Top             =   300
               Width           =   1455
            End
         End
      End
      Begin VB.HScrollBar hsVsl 
         Height          =   270
         Left            =   120
         SmallChange     =   20
         TabIndex        =   23
         Top             =   5760
         Width           =   4965
      End
      Begin VB.VScrollBar vsVsl 
         Height          =   3540
         Left            =   9120
         SmallChange     =   20
         TabIndex        =   24
         Top             =   840
         Width           =   270
      End
      Begin VB.PictureBox picVessel 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   5460
         Left            =   0
         ScaleHeight     =   364
         ScaleMode       =   3  '픽셀
         ScaleWidth      =   600
         TabIndex        =   25
         Top             =   0
         Width           =   9000
         Begin VB.PictureBox picHatch 
            Appearance      =   0  '평면
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H80000008&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2415
            Index           =   0
            Left            =   3120
            ScaleHeight     =   159
            ScaleMode       =   3  '픽셀
            ScaleWidth      =   255
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
            Width           =   3855
            Begin VB.PictureBox Pic 
               Appearance      =   0  '평면
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  '없음
               FillColor       =   &H80000008&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1710
               Index           =   0
               Left            =   600
               ScaleHeight     =   114
               ScaleMode       =   3  '픽셀
               ScaleWidth      =   122
               TabIndex        =   30
               Top             =   120
               Visible         =   0   'False
               Width           =   1830
               Begin VB.TextBox txtNStkWgt 
                  Alignment       =   2  '가운데 맞춤
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  '없음
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   270
                  Index           =   0
                  Left            =   1440
                  MaxLength       =   4
                  TabIndex        =   61
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.TextBox txtSNo 
                  Alignment       =   2  '가운데 맞춤
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  '없음
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   270
                  Index           =   0
                  Left            =   120
                  MaxLength       =   2
                  TabIndex        =   31
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Shape shpBox 
                  BorderColor     =   &H000000C0&
                  BorderStyle     =   3  '점
                  Height          =   495
                  Index           =   0
                  Left            =   240
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1335
               End
            End
         End
         Begin VB.Image imgCenter 
            Height          =   300
            Left            =   1440
            Picture         =   "frmVesExplorer.frx":1289
            Top             =   4800
            Width           =   19575
         End
         Begin VB.Image imgBridge 
            Height          =   1500
            Left            =   1440
            Picture         =   "frmVesExplorer.frx":35BA
            Top             =   2160
            Width           =   690
         End
         Begin VB.Image imgRight 
            Height          =   2580
            Left            =   7920
            Picture         =   "frmVesExplorer.frx":4F64
            Top             =   2520
            Width           =   1155
         End
         Begin VB.Image imgLeft 
            Height          =   2580
            Left            =   240
            Picture         =   "frmVesExplorer.frx":6CD1
            Top             =   2520
            Width           =   1155
         End
      End
   End
   Begin VB.PictureBox picBackCenter 
      BorderStyle     =   0  '없음
      Height          =   6555
      Left            =   2760
      MousePointer    =   9  'W E 크기 조정
      ScaleHeight     =   437
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   19
      TabIndex        =   21
      Top             =   360
      Width           =   285
      Begin VB.PictureBox picCenter 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  '없음
         Height          =   2715
         Left            =   90
         ScaleHeight     =   181
         ScaleMode       =   3  '픽셀
         ScaleWidth      =   7
         TabIndex        =   22
         Top             =   3840
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2040
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":8980
            Key             =   "imgHide"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":8F1A
            Key             =   "imgExpand"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":9074
            Key             =   "imgStruct"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":960E
            Key             =   "imgVessel"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":99A8
            Key             =   "imgHatch"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":9F42
            Key             =   "imgBay"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":A4DC
            Key             =   "imgTable"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":AA76
            Key             =   "imgBasic"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":B010
            Key             =   "imgClose"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":B4CA
            Key             =   "imgZoomOut"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":BA64
            Key             =   "imgZoomIn"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":BFFE
            Key             =   "imgDraw"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":C598
            Key             =   "ImgCellGuide"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":CB32
            Key             =   "imgHOpen"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":D0CC
            Key             =   "imgHFold"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":D666
            Key             =   "imgCopy"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":DC00
            Key             =   "imgPaste"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":E19A
            Key             =   "imgCopyCheck"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":E734
            Key             =   "imgRefresh"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVesExplorer.frx":ECCE
            Key             =   "imgZoom"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   13830
      _ExtentX        =   24395
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   33
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgHide"
            Object.ToolTipText     =   "Hide Vessel Structure List"
            ImageKey        =   "imgHide"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgExpand"
            Object.ToolTipText     =   "Collapse Vessel Structure List"
            ImageKey        =   "imgExpand"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgClose"
            Object.ToolTipText     =   "Close Window"
            ImageKey        =   "imgClose"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgRefresh"
            Object.ToolTipText     =   "Reset Vessel Structure"
            ImageKey        =   "imgRefresh"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgDraw"
            Object.ToolTipText     =   "Draw Slot Allocation"
            ImageKey        =   "imgDraw"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ImgCellGuide"
            Object.ToolTipText     =   "Slim Cell Guide"
            ImageKey        =   "ImgCellGuide"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgHOpen"
            Object.ToolTipText     =   "Open Hatch Cover"
            ImageKey        =   "imgHOpen"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgHFold"
            Object.ToolTipText     =   "Folding Hatch Cover"
            ImageKey        =   "imgHFold"
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgZoomOut"
            Object.ToolTipText     =   "Zoom Out"
            ImageKey        =   "imgZoomOut"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgZoom"
            Object.ToolTipText     =   "Default Size"
            ImageKey        =   "imgZoom"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgZoomIn"
            Object.ToolTipText     =   "Zoom In"
            ImageKey        =   "imgZoomIn"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgCopyCheck"
            Object.ToolTipText     =   "Copy Condition"
            ImageKey        =   "imgCopyCheck"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "chkHatch"
                  Text            =   "Hattch"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "chkBay"
                  Text            =   "Bay"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "chkCover"
                  Text            =   "Cover"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "chkDeck"
                  Text            =   "Deck"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "chkHold"
                  Text            =   "Hold"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "chkDeckNo"
                  Text            =   "Deck No."
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "chkHoldNo"
                  Text            =   "Hold No."
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgCopy"
            Object.ToolTipText     =   "Copy Vessel Structure"
            ImageKey        =   "imgCopy"
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "imgPaste"
            Object.ToolTipText     =   "Paste Vessel Structure"
            ImageKey        =   "imgPaste"
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.OptionButton optBay 
         Caption         =   "Bay"
         Height          =   180
         Left            =   11760
         TabIndex        =   57
         Top             =   80
         Width           =   735
      End
      Begin VB.OptionButton optHatch 
         Caption         =   "Hatch"
         Height          =   180
         Left            =   10920
         TabIndex        =   56
         Top             =   80
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '없음
         Height          =   180
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Synchro"
         Top             =   80
         Width           =   705
      End
      Begin VB.TextBox txtCopy 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '없음
         Height          =   180
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Copy Bay"
         Top             =   80
         Width           =   1290
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '없음
         Height          =   180
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "HC Type"
         Top             =   80
         Width           =   930
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '없음
         Height          =   180
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Draw Mode"
         Top             =   80
         Width           =   960
      End
      Begin VB.TextBox txtZoom 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '없음
         Height          =   180
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "150 %"
         Top             =   80
         Width           =   480
      End
   End
   Begin MSComctlLib.TreeView tvList 
      Height          =   6945
      Left            =   0
      TabIndex        =   18
      Top             =   360
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   12250
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmVesExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cDX As Single           '스크롤바관련
Dim bFlagHide As Boolean    'TREE VIEW 숨기기
Dim bFlagExpand As Boolean  'TREE VIEW 펼치기

Dim gDrawMode As String     'A:All, H:Hatch, B:Bay
Dim gDrawNo As String
Dim gCopyMode As String     'H:Hatch,B:Bay,D:Deck,HD:Hold,DN:Deck No,HN:Hold No, HC:Hatch Cover
Dim gSynchro As String      'H:Hatch, B:Bay

Dim gPicW%, gPicH%          '픽쳐박스 크기

Dim gBayNoH As Single
Dim gYcW As Single, gYcH As Single                                  '셀크기
Dim gHCoverW As Single, gHCoverH As Single                          '해치커버크기
Dim gMarginL As Single, gMarginT As Single, gHCoverGap As Single    '간격
Dim gCellGapW As Single, gCellGapH As Single                        '셀간격
Dim gZoomIdx As Single                                              '줌인,아웃 퍼센트
Dim gBayFondSize As Single, gNoFontSize As Single, gWgtFontSize As Single
Const gHatchCoverVal = 0.25

Dim gDrawCell As String         'Draw Cell 모드 C:cell, G:slim cell guide
Dim gBColor As Long
Dim gBWgtColor As Long
Dim gCellGuideColor As Long
Dim gDrawHCover As String         'Draw Hatch Cover 모드 O:Open, F:Fold
Dim gHcOpenColor As Long
Dim gHcFoldColor As Long

'픽쳐박스 마우스 이벤트
Dim gStartX%, gStartY%

'Row NO, Tier NO
Dim gRowTier As String
Dim gHatchIdx%, gBayIdx%, gHdIdx%, gSnoIdx%

'Row Stack Weight
Dim gNStkWgtIdx%

'copy
Dim gCopyHatchIdx%, gCopyBayIdx%

Private Sub cmdChage_Click()
  Dim i%, j%, k%, l%, hFlag As Boolean
  Dim iVal%, iVal2%, idx%
  Dim iOrgNoHatchs%, iOrgNoBays%
  Dim iOrgMaxRows%, iOrgMaxTiers%
  
  hFlag = False
  
  If MsgBox("Are you sure to change? It effects basic structure", vbQuestion + vbYesNo, "질문") = vbNo Then Exit Sub
  
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
    
    .iDMaxRows = Val(txtDeckRows.Text)
    .iHMaxRows = Val(txtHoldRows.Text)
    .iDMaxTiers = Val(txtDeckTiers.Text)
    .iHMaxTiers = Val(txtHoldTiers.Text)
    
    .iBgNo = Val(txtBgNo.Text)
    .nBgLength = Val(txtBgLength.Text)
    .bSaveFlag = True
    
    iOrgMaxRows = .iMaxRows
    iOrgMaxTiers = .iMaxTiers
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
    
    If gVessel.iNoHatchs < Val(txtHatchCnt.Text) Then
      hFlag = True
      iOrgNoHatchs = gVessel.iNoHatchs
      iOrgNoBays = gVessel.iNoBays
      
      .iNoHatchs = Val(txtHatchCnt.Text)
      .iNoBays = .iNoBays + (.iNoHatchs - iOrgNoHatchs) * 3
    End If
    
    If hFlag Then
      'Hatch Info
      ReDim Preserve .gHatch(1 To .iNoHatchs)
      For i = iOrgNoHatchs To .iNoHatchs
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
    End If
    
    'Bay Info
    ReDim Preserve .gBay(1 To .iNoBays)
    For i = 1 To .iNoBays
      .gBay(i).bSaveFlag = True
      .gBay(i).iNoRows(0) = .iHMaxRows
      .gBay(i).iNoRows(1) = .iDMaxRows
      .gBay(i).iNoTiers(0) = .iHMaxTiers
      .gBay(i).iNoTiers(1) = .iDMaxTiers
      
      If i = iOrgNoBays Then idx = Val(.gBay(i).sBayNo)
      'Hatch 개수 변경 시
      If hFlag And i > iOrgNoBays Then
        .gBay(i).iBay = i
        .gBay(i).iHchNo = iOrgNoHatchs + Int((i - iOrgNoBays - 1) / 3) + 1
        
        idx = idx + 1
        If (i - iOrgNoBays + 1) Mod 3 = 0 Then
          .gBay(i).sSize = "4"
          If idx Mod 2 = 1 Then idx = idx + 1
        Else
          .gBay(i).sSize = "2"
          If idx Mod 2 = 0 Then idx = idx + 1
        End If
        
        .gBay(i).sBayNo = Format(idx, "00")
        
        
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
        
      Else
        ReDim Preserve .gBay(i).gRow(0 To 1, 1 To .iMaxRows)
        ReDim Preserve .gBay(i).gTier(0 To 1, 1 To .iMaxTiers)
        
        ReDim .gBay(i).gCell(0 To 1, 1 To .iMaxRows, 1 To .iMaxTiers)
        For j = 0 To 1
          Call PasteCellInfo(i, i, j, "C", iOrgMaxRows, iOrgMaxTiers)
        Next j
        
        ReDim .gBay(i).gCellOrg(0 To 1, 1 To .iMaxRows, 1 To .iMaxTiers)
        For j = 0 To 1
          Call PasteCellInfo(i, i, j, "O", .iMaxRows, .iMaxTiers)
        Next j
        
      End If
    Next i
  End With
  
  '트리 초기화
  Call SetTreeViewItem
  mdiVesConfig.Caption = "Vessel Configuration (v2.0) - " & gVessel.sVCode
End Sub

Private Sub Form_Load()
  Me.Height = 9165
  Me.Width = 14890
  
  'TREE VIEW 숨기기, 펼치기 변수
  bFlagHide = False
  bFlagExpand = True
  
  gZoomIdx = 1.5
  gDrawCell = "C"
  gBColor = RGB(70, 65, 217)
  gBWgtColor = RGB(219, 0, 0)
  gCellGuideColor = vbRed
  gDrawHCover = "O"
  gHcOpenColor = RGB(242, 150, 97)
  gHcFoldColor = RGB(0, 216, 155)
  
  gCopyMode = "B"
  gSynchro = "H"
  
  '트리 초기화
  Call SetTreeViewItem
  'Vessel View
  gDrawMode = "V"
  Call SetDrawVessel
  Call ReSizePic
End Sub

Private Sub SetSize()
  Select Case gDrawMode
    Case "A"
      gYcW = 15 * gZoomIdx: gYcH = 12.5 * gZoomIdx
    Case "H"
      gYcW = 18 * gZoomIdx: gYcH = 12.5 * gZoomIdx
    Case "B"
      gYcW = 30 * gZoomIdx: gYcH = 20 * gZoomIdx
  End Select
  
  gCellGapW = gYcW / 7: gCellGapH = gYcH / 7
  gHCoverW = (gYcW + gCellGapW) / 4: gHCoverH = gHCoverW
  gHCoverGap = gCellGapH * 1.5
  gMarginL = gYcW * 1.5: gMarginT = gYcH
  gBayNoH = gYcH * 2
  gBayFondSize = gYcW: gNoFontSize = gBayFondSize / 2.5: gWgtFontSize = gBayFondSize / 3.5
End Sub

Private Sub Form_Resize()
  Dim MX%, MY%, MW%, MH%
  
  MX = 0
  MY = (tbrMain.Height + 2)
  MW = 150
  MH = Me.ScaleHeight - (tbrMain.Height + 2) - 2
  If MH < 0 Then Exit Sub
  tvList.Move MX, MY, MW, MH
  
  MX = tvList.Width
  MY = tvList.Top
  MW = 5
  MH = tvList.Height
  picBackCenter.Move MX, MY, MW, MH
  
  If bFlagHide Then
    MX = 0
    MY = tvList.Top
    MW = Me.ScaleWidth - 2
    MH = tvList.Height
  Else
    MX = picBackCenter.Left + picBackCenter.Width
    MY = tvList.Top
    MW = Me.ScaleWidth - tvList.Width - picBackCenter.Width - 2
    MH = tvList.Height
  End If
  picBack.Move MX, MY, MW, MH
  
  fraTable.Move 0, 0, picBack.Width, picBack.Height
  fraBasic.Move 0, 0, picBack.Width, picBack.Height
  
  Call SetScroll
End Sub

Private Sub SetScroll()
  Dim vFlag As Boolean
  
  vsVsl.Move picBack.ScaleWidth - vsVsl.Width, 0, vsVsl.Width, picBack.ScaleHeight - hsVsl.Height
  hsVsl.Move 0, picBack.ScaleHeight - hsVsl.Height, picBack.ScaleWidth - vsVsl.Width, hsVsl.Height
  
  vFlag = False
  If picBack.ScaleHeight < picVessel.ScaleHeight Then
    vsVsl.Visible = True
    vsVsl.Max = picVessel.Height - picBack.ScaleHeight
    vsVsl.LargeChange = vsVsl.Max
    vsVsl.Value = 0
    vFlag = True
  Else
    vsVsl.Visible = False
  End If
    
  If vFlag = False Then
    vsVsl.Max = 0: vsVsl.LargeChange = 1
  End If
  
  vFlag = False
  If picBack.ScaleWidth < picVessel.ScaleWidth Then
    hsVsl.Visible = True
    hsVsl.Max = picVessel.Width - picBack.ScaleWidth
    hsVsl.LargeChange = hsVsl.Max
    vFlag = True
  Else
    hsVsl.Visible = False
  End If
  
  If vFlag = False Then
    hsVsl.Max = 0: hsVsl.LargeChange = 1
  End If
End Sub

Private Sub SetDrawVessel()
  Const iGap = 10: Const iMargin = 2: Const iWid = 20: Const iHei = 10
  Const iTop = 70
  Dim i%, j%, k%, l%, iVal%, iTierIdx%
  Dim iStartX%, iStartY%, xPos%, yPos%
  Dim iStd%, iBX%, iBY%, tStr$
  Dim sBayNo$, iMaxRows%, bFillFlag As Boolean
  
  picVessel.Width = 1200: picVessel.Height = 600
  
  '측면
  iStartX = iGap + imgLeft.Width + iMargin
  iStartY = iGap + iTop
  
  picVessel.DrawWidth = 1
  picVessel.FontSize = 8
  picVessel.ForeColor = RGB(150, 150, 150)
  iVal = 0
  For i = gVessel.iNoBays To 1 Step -1
    If iVal <> gVessel.gBay(i).iHchNo Then
      If iVal <> 0 Then iStartX = iStartX + iMargin
      iVal = gVessel.gBay(i).iHchNo
      
      If gVessel.iBgNo = iVal Then
        iBX = iStartX
        
        iStartX = iStartX + imgBridge.Width + iMargin
      End If
      
      tStr = gVessel.gBay(i).sSize
    End If
    
    If gVessel.gBay(i).sSize = tStr Then
      For j = 1 To gVessel.iDMaxTiers + gVessel.iHMaxTiers
        xPos = iStartX
        yPos = iStartY + (j - 1) * iHei
        
        bFillFlag = False
        If j <= gVessel.iDMaxTiers Then
          iTierIdx = gVessel.iDMaxTiers - (j - 1)
          For k = 1 To gVessel.gBay(i).iNoRows(1)
            If gVessel.gBay(i).gCell(1, k, iTierIdx).sSS = "Y" Then
              bFillFlag = True
              Exit For
            End If
          Next k
        Else
          iTierIdx = gVessel.iHMaxTiers - (j - gVessel.iDMaxTiers - 1)
          For k = 1 To gVessel.gBay(i).iNoRows(0)
            If gVessel.gBay(i).gCell(0, k, iTierIdx).sSS = "Y" Then
              bFillFlag = True
              Exit For
            End If
          Next k
        End If
        
        If bFillFlag Then
          picVessel.Line (iStartX, yPos)-Step(iWid * (Val(tStr) / 2), iHei), vbBlue, BF
        End If
        picVessel.Line (iStartX, yPos)-Step(iWid * (Val(tStr) / 2), iHei), RGB(150, 150, 150), B
        
        If j = gVessel.iDMaxTiers + 1 And iStd = 0 Then
          iStd = yPos
          iBY = iStd - imgBridge.Height
        End If
      Next j
      
      iStartX = iStartX + iWid * (Val(tStr) / 2)
    End If
    
    sBayNo = gVessel.gBay(i).sBayNo
    If gVessel.gBay(i).sSize = "2" Then
      picVessel.CurrentX = xPos + (iWid - picVessel.TextWidth(sBayNo)) / 2
      picVessel.CurrentY = iStartY - 30
    Else
      picVessel.CurrentX = xPos + (iWid * 2 - picVessel.TextWidth(sBayNo)) / 2
      picVessel.CurrentY = iStartY - 40
    End If
    picVessel.Print sBayNo
  Next i
  
  'Image
  imgLeft.Left = iGap
  imgLeft.Top = iStd - 80
  imgCenter.Left = imgLeft.Left + imgLeft.Width
  imgCenter.Top = imgLeft.Top + imgLeft.Height - imgCenter.Height
  imgCenter.Width = iStartX - (iGap + imgLeft.Width + iMargin) + 5
  imgRight.Left = imgCenter.Left + imgCenter.Width
  imgRight.Top = imgLeft.Top
  imgBridge.Left = iBX
  imgBridge.Top = iBY
  
  '상면
  iStartX = iGap + imgLeft.Width + iMargin
  iStartY = imgRight.Top + imgRight.Height + iTop
  iVal = 0
  For i = gVessel.iNoBays To 1 Step -1
    If iVal <> gVessel.gBay(i).iHchNo Then
      If iVal <> 0 Then iStartX = iStartX + iMargin
      iVal = gVessel.gBay(i).iHchNo

      If gVessel.iBgNo = iVal Then
        iBX = iStartX

        iStartX = iStartX + imgBridge.Width + iMargin
      End If

      tStr = gVessel.gBay(i).sSize
    End If

    If gVessel.gBay(i).sSize = tStr Then
      If gVessel.iDMaxRows > gVessel.iHMaxRows Then
        iMaxRows = gVessel.iDMaxRows
      Else
        iMaxRows = gVessel.iHMaxRows
      End If
      
      For j = 1 To iMaxRows
        xPos = iStartX
        yPos = iStartY + (j - 1) * iHei
        
        bFillFlag = False
        For k = 0 To 1
          For l = 1 To gVessel.gBay(i).iNoTiers(k)
            If gVessel.gBay(i).gCell(k, j, l).sSS = "Y" Then
              bFillFlag = True
              Exit For
            End If
          Next l
          
          If bFillFlag Then Exit For
        Next k
        
        If bFillFlag Then
          picVessel.Line (iStartX, yPos)-Step(iWid * (Val(tStr) / 2), iHei), vbBlue, BF
        End If
        picVessel.Line (iStartX, yPos)-Step(iWid * (Val(tStr) / 2), iHei), RGB(150, 150, 150), B
      Next j

      iStartX = iStartX + iWid * (Val(tStr) / 2)
    End If
  Next i
  
  picVessel.FillStyle = 7
  picVessel.Line (imgLeft.Left, iStartY)-Step(imgLeft.Width, iMaxRows * iHei), vbBlack, B
  picVessel.FillStyle = 4
  picVessel.Line (imgRight.Left, iStartY)-Step(imgRight.Width, iMaxRows * iHei / 2), vbBlack, B
  picVessel.FillStyle = 5
  picVessel.Line (imgRight.Left, iStartY + iMaxRows * iHei / 2)-Step(imgRight.Width, iMaxRows * iHei / 2), vbBlack, B
  picVessel.FillStyle = 1
  picVessel.Line (imgBridge.Left, iStartY)-Step(imgBridge.Width, iMaxRows * iHei), vbBlack, B
  picVessel.Line (imgBridge.Left + iGap / 2, iStartY + iGap / 2)-Step(imgBridge.Width - iGap, iMaxRows * iHei - iGap), RGB(120, 120, 120), BF
  
  gPicW = imgRight.Left + imgRight.Width + iGap * 5
  gPicH = yPos + iGap * 5
End Sub

Private Function GetBayIdx(sBayNo$) As Integer
  Dim i%
  
  GetBayIdx = 0
  For i = 1 To gVessel.iNoBays
    If gVessel.gBay(i).sBayNo = sBayNo Then
      GetBayIdx = i
      Exit For
    End If
  Next i
End Function

Private Function GetHatchIdx(sBayNo$) As Integer
  Dim i%
  
  GetHatchIdx = 0
  For i = 1 To gVessel.iNoBays
    If gVessel.gBay(i).sBayNo = sBayNo Then
      GetHatchIdx = gVessel.gBay(i).iHchNo
      Exit For
    End If
  Next i
End Function

Private Sub SetDrawStructure()
  Dim i%
  Dim iPage%
  
  If gVessel.iNoHatchs Mod 2 = 1 Then
    iPage = (gVessel.iNoHatchs + 1) / 2
  Else
    iPage = gVessel.iNoHatchs / 2
  End If
  
  
  For i = 1 To gVessel.iNoHatchs
    Call SetDrawHatch(Format(i, "00"))
    
    If i > iPage Then
      picHatch(i - 1).Left = (iPage - (i - iPage)) * picHatch(i - 1).Width
      picHatch(i - 1).Top = picHatch(i - 1).Height
    Else
      picHatch(i - 1).Left = (iPage - i) * picHatch(i - 1).Width
      picHatch(i - 1).Top = 0
    End If
    
    picHatch(i - 1).BorderStyle = 1
  Next i
  
  gPicW = picHatch(0).Width * iPage: gPicH = picHatch(0).Height * 2
End Sub

Private Sub SetDrawHatch(sHch$)
  Dim i%, idx%, iHch%
  Dim tStr$
  
  idx = 0
  iHch = Val(sHch)
  picHatch(iHch - 1).Cls
  picHatch(iHch - 1).Visible = True
  picHatch(iHch - 1).BorderStyle = 0
  For i = 1 To gVessel.iNoBays
    If gVessel.gBay(i).iHchNo = iHch Then
      idx = idx + 1
      Call DrawBay(gVessel.gBay(i).iBay)
      Select Case idx
        Case 1
          picHatch(iHch - 1).Move 0, 0, Pic(gVessel.gBay(i).iBay - 1).Width * 2, Pic(gVessel.gBay(i).iBay - 1).Height * 2
          gPicW = picHatch(iHch - 1).Width: gPicH = picHatch(iHch - 1).Height
          
          If gVessel.gBay(i).sSize = "4" Then
            Pic(gVessel.gBay(i).iBay - 1).Left = 0
            Pic(gVessel.gBay(i).iBay - 1).Top = Pic(gVessel.gBay(i).iBay - 1).Height / 2
          Else
            Pic(gVessel.gBay(i).iBay - 1).Left = Pic(gVessel.gBay(i).iBay - 1).Width
            Pic(gVessel.gBay(i).iBay - 1).Top = 0
          End If
          
        Case 2
          Pic(gVessel.gBay(i).iBay - 1).Left = Pic(gVessel.gBay(i).iBay - 1).Width
          Pic(gVessel.gBay(i).iBay - 1).Top = Pic(gVessel.gBay(i).iBay - 1).Height
          
        Case 3
          Pic(gVessel.gBay(i - 1).iBay - 1).Left = 0
          Pic(gVessel.gBay(i - 1).iBay - 1).Top = Pic(gVessel.gBay(i - 1).iBay - 1).Height / 2
          
          Pic(gVessel.gBay(i).iBay - 1).Left = Pic(gVessel.gBay(i).iBay - 1).Width
          Pic(gVessel.gBay(i).iBay - 1).Top = Pic(gVessel.gBay(i).iBay - 1).Height
          
      End Select
    End If
  Next i
  
  'hatch no.
  picHatch(iHch - 1).FontBold = False
  picHatch(iHch - 1).FontName = "HY견명조"
  picHatch(iHch - 1).FontSize = gBayFondSize
  picHatch(iHch - 1).ForeColor = vbBlack
  tStr = "HATCH " & sHch
  picHatch(iHch - 1).CurrentX = (picHatch(iHch - 1).Width / 2 - picHatch(iHch - 1).TextWidth(tStr)) / 2
  picHatch(iHch - 1).CurrentY = gMarginT + (gBayNoH - picHatch(iHch - 1).TextHeight(tStr)) / 2
  picHatch(iHch - 1).Print tStr
End Sub

Private Sub SetDrawBay(sBayNo$)
  Dim iBayIdx%
  
  iBayIdx = GetBayIdx(sBayNo)
  If iBayIdx = 0 Then Exit Sub
  Call DrawBay(iBayIdx)
  
  picHatch(GetHatchIdx(sBayNo) - 1).Visible = True
  picHatch(GetHatchIdx(sBayNo) - 1).BorderStyle = 0
  picHatch(GetHatchIdx(sBayNo) - 1).Move 0, 0, Pic(iBayIdx - 1).Width, Pic(iBayIdx - 1).Height
  
  gPicW = Pic(iBayIdx - 1).Width: gPicH = Pic(iBayIdx - 1).Height
End Sub

Private Sub DrawBay(iBayIdx%)
  Dim i%, j%, iTierIdx%
  Dim iHcF%, iHcT%
  Dim idxP%, tStr$
  Dim iStartX As Single, iStartY As Single, xPos As Single, yPos As Single
  Dim iMaxRows%, iMaxTiers%
  
  Call SetSize
  
  idxP = iBayIdx - 1
  Pic(idxP).Cls
  
  'PicBox Size
  iMaxRows = gVessel.iMaxRows
  iMaxTiers = gVessel.iMaxTiers
  Pic(idxP).Left = 0: Pic(idxP).Top = 0
  Pic(idxP).Width = gMarginL + (iMaxRows + 1) * (gYcW + gCellGapW)
  Pic(idxP).Height = gMarginT + gBayNoH + (gVessel.iDMaxTiers + gVessel.iHMaxTiers + 2 + 2) * (gYcH + gCellGapH) + (gHCoverGap + gHCoverH * 2)
  
  
  Pic(idxP).ForeColor = gBColor
  'BayNo
  Pic(idxP).FontBold = False
  Pic(idxP).FontName = "HY견명조"
  Pic(idxP).FontSize = gBayFondSize
  tStr = "BAY " & gVessel.gBay(iBayIdx).sBayNo
  Pic(idxP).CurrentX = gMarginL + (iMaxRows * (gYcW + gCellGapW) - Pic(idxP).TextWidth(tStr)) / 2
  Pic(idxP).CurrentY = gMarginT + (gBayNoH - Pic(idxP).TextHeight(tStr)) / 2
  Pic(idxP).Print tStr
  
  'Row Stacking Weight
  Pic(idxP).ForeColor = gBWgtColor
  Pic(idxP).FontName = "Arial"
  Pic(idxP).FontSize = gWgtFontSize
  For i = 0 To 1
    For j = 1 To gVessel.gBay(iBayIdx).iNoRows(i)
      tStr = Trim(Str(gVessel.gBay(iBayIdx).gRow(i, j).nStkWgt))
      tStr = Format(tStr, "0000")
      If i = 0 Then
        Pic(idxP).CurrentX = gMarginL + (iMaxRows * (gYcW + gCellGapW) - gVessel.iHMaxRows * (gYcW + gCellGapW)) / 2 + (gYcW + gCellGapW) * (j - 1) + (gYcW - Pic(idxP).TextWidth(tStr)) / 2
        Pic(idxP).CurrentY = gMarginT + gBayNoH + (gVessel.iDMaxTiers + gVessel.iHMaxTiers + 3) * (gYcH + gCellGapH) + (gHCoverGap + gHCoverH * 2) + (gYcH + gCellGapH - Pic(idxP).TextHeight(tStr)) / 2
      Else
        Pic(idxP).CurrentX = gMarginL + (iMaxRows * (gYcW + gCellGapW) - gVessel.iDMaxRows * (gYcW + gCellGapW)) / 2 + (gYcW + gCellGapW) * (j - 1) + (gYcW - Pic(idxP).TextWidth(tStr)) / 2
        Pic(idxP).CurrentY = gMarginT + gBayNoH + (gYcH + gCellGapH - Pic(idxP).TextHeight(tStr)) / 2
      End If
      
      gVessel.gBay(iBayIdx).gRow(i, j).iL_Wgt = Pic(idxP).CurrentX
      gVessel.gBay(iBayIdx).gRow(i, j).iT_Wgt = Pic(idxP).CurrentY
      gVessel.gBay(iBayIdx).gRow(i, j).iW_Wgt = Pic(idxP).TextWidth(tStr)
      gVessel.gBay(iBayIdx).gRow(i, j).iH_Wgt = Pic(idxP).TextHeight(tStr)
      
      Select Case Len(Trim(Str(gVessel.gBay(iBayIdx).gRow(i, j).nStkWgt)))
        Case 1
          Pic(idxP).CurrentX = Pic(idxP).CurrentX + Pic(idxP).TextWidth("000") / 2
        Case 2
          Pic(idxP).CurrentX = Pic(idxP).CurrentX + Pic(idxP).TextWidth("00") / 2
        Case 3
          Pic(idxP).CurrentX = Pic(idxP).CurrentX + Pic(idxP).TextWidth("0") / 2
      End Select
      
      If gVessel.gBay(iBayIdx).gRow(i, j).nStkWgt > 0 Then Pic(idxP).Print Trim(Str(gVessel.gBay(iBayIdx).gRow(i, j).nStkWgt))
    Next j
  Next i
  
  'Row NO
  Pic(idxP).ForeColor = gBColor
  Pic(idxP).FontSize = gNoFontSize
  Pic(idxP).FontBold = True
  For i = 0 To 1
    For j = 1 To gVessel.gBay(iBayIdx).iNoRows(i)
      tStr = gVessel.gBay(iBayIdx).gRow(i, j).sNo
      If tStr = "" Then tStr = "00"
      
      If i = 0 Then
        Pic(idxP).CurrentX = gMarginL + (iMaxRows * (gYcW + gCellGapW) - gVessel.iHMaxRows * (gYcW + gCellGapW)) / 2 + (gYcW + gCellGapW) * (j - 1) + (gYcW - Pic(idxP).TextWidth(tStr)) / 2
        Pic(idxP).CurrentY = gMarginT + gBayNoH + (gVessel.iDMaxTiers + gVessel.iHMaxTiers + 1 + 1) * (gYcH + gCellGapH) + (gHCoverGap + gHCoverH * 2) + (gYcH + gCellGapH - Pic(idxP).TextHeight(tStr)) / 2
      Else
        Pic(idxP).CurrentX = gMarginL + (iMaxRows * (gYcW + gCellGapW) - gVessel.iDMaxRows * (gYcW + gCellGapW)) / 2 + (gYcW + gCellGapW) * (j - 1) + (gYcW - Pic(idxP).TextWidth(tStr)) / 2
        Pic(idxP).CurrentY = gMarginT + gBayNoH + (gYcH + gCellGapH) + (gYcH + gCellGapH - Pic(idxP).TextHeight(tStr)) / 2
      End If
      
      gVessel.gBay(iBayIdx).gRow(i, j).iL = Pic(idxP).CurrentX
      gVessel.gBay(iBayIdx).gRow(i, j).iT = Pic(idxP).CurrentY
      gVessel.gBay(iBayIdx).gRow(i, j).iW = Pic(idxP).TextWidth(tStr)
      gVessel.gBay(iBayIdx).gRow(i, j).iH = Pic(idxP).TextHeight(tStr)
      
      Pic(idxP).Print gVessel.gBay(iBayIdx).gRow(i, j).sNo
    Next j
  Next i
  
  
  'Tier No
  iStartX = gMarginL + iMaxRows * (gYcW + gCellGapW) + (gYcW + gCellGapW) / 2
  For i = 0 To 1
    For j = 1 To gVessel.gBay(iBayIdx).iNoTiers(i)
      iTierIdx = gVessel.gBay(iBayIdx).iNoTiers(i) - (j - 1)
      
      tStr = gVessel.gBay(iBayIdx).gTier(i, iTierIdx).sNo
      If tStr = "" Then tStr = "00"
      
      Pic(idxP).CurrentX = iStartX - Pic(idxP).TextWidth(tStr) / 2
      If i = 0 Then
        Pic(idxP).CurrentY = gMarginT + gBayNoH + (gVessel.iDMaxTiers + 1) * (gYcH + gCellGapH) + (gHCoverGap + gHCoverH * 2) + (gYcH + gCellGapH - Pic(idxP).TextHeight(tStr)) / 2
      Else
        Pic(idxP).CurrentY = gMarginT + gBayNoH + (gYcH + gCellGapH) + (gYcH + gCellGapH - Pic(idxP).TextHeight(tStr)) / 2
      End If
      Pic(idxP).CurrentY = Pic(idxP).CurrentY + (gYcH + gCellGapH) * (j - 1 + 1)
      
      gVessel.gBay(iBayIdx).gTier(i, iTierIdx).iL = Pic(idxP).CurrentX
      gVessel.gBay(iBayIdx).gTier(i, iTierIdx).iT = Pic(idxP).CurrentY
      gVessel.gBay(iBayIdx).gTier(i, iTierIdx).iW = Pic(idxP).TextWidth(tStr)
      gVessel.gBay(iBayIdx).gTier(i, iTierIdx).iH = Pic(idxP).TextHeight(tStr)
      
      Pic(idxP).Print gVessel.gBay(iBayIdx).gTier(i, iTierIdx).sNo
    Next j
  Next i
  
  'Deck 그리기
  Pic(idxP).DrawWidth = 1
  iStartX = gMarginL + (iMaxRows * (gYcW + gCellGapW) - gVessel.iDMaxRows * (gYcW + gCellGapW)) / 2
  iStartY = gMarginT + gBayNoH + (gYcH + gCellGapH) * 2
  For i = 1 To gVessel.iDMaxRows
    xPos = iStartX + (i - 1) * (gYcW + gCellGapW)
    For j = 1 To gVessel.iDMaxTiers
      iTierIdx = gVessel.iDMaxTiers - (j - 1)
      yPos = iStartY + (j - 1) * (gYcH + gCellGapH)
      
      If gVessel.gBay(iBayIdx).gCell(1, i, iTierIdx).sSS = "Y" Then
        Pic(idxP).Line (xPos, yPos)-Step(gYcW, gYcH), gBColor, BF
      End If
      Pic(idxP).Line (xPos, yPos)-Step(gYcW, gYcH), vbBlack, B
      
      gVessel.gBay(iBayIdx).gCell(1, i, iTierIdx).iL = xPos
      gVessel.gBay(iBayIdx).gCell(1, i, iTierIdx).iT = yPos
      gVessel.gBay(iBayIdx).gCell(1, i, iTierIdx).iW = gYcW
      gVessel.gBay(iBayIdx).gCell(1, i, iTierIdx).iH = gYcH
      
    Next j
  Next i
  
  'Hatch Cover 그리기
  Pic(idxP).DrawWidth = 0.8
  iStartX = gMarginL - gCellGapW / 2
  iStartY = gMarginT + gBayNoH + (gVessel.iDMaxTiers + 1 + 1) * (gYcH + gCellGapH) + gHCoverGap
  yPos = iStartY
'  For i = 1 To gVessel.gHatch(gVessel.gBay(iBayIdx).iHchNo).iNoCvr
'    iHcF = gVessel.gHatch(gVessel.gBay(iBayIdx).iHchNo).gHchCvr(i).nF / 0.25 + 1
'    iHcT = gVessel.gHatch(gVessel.gBay(iBayIdx).iHchNo).gHchCvr(i).nT / 0.25
'
'    For j = iHcF To iHcT
'      xPos = iStartX + (j - 1) * gHCoverW
'      If gVessel.gHatch(gVessel.gBay(iBayIdx).iHchNo).gHchCvr(i).sHType = "O" Then
'        Pic(idxP).Line (xPos, yPos)-Step(gHCoverW, gHCoverH), gHcOpenColor, BF
'      Else
'        Pic(idxP).Line (xPos, yPos)-Step(gHCoverW, gHCoverH), gHcFoldColor, BF
'      End If
'    Next j
'  Next i
  For i = 1 To gVessel.gBay(iBayIdx).iNoCvr
    iHcF = gVessel.gBay(iBayIdx).gHchCvr(i).nF / 0.25 + 1
    iHcT = gVessel.gBay(iBayIdx).gHchCvr(i).nT / 0.25
    
    For j = iHcF To iHcT
      xPos = iStartX + (j - 1) * gHCoverW
      If gVessel.gBay(iBayIdx).gHchCvr(i).sHType = "O" Then
        Pic(idxP).Line (xPos, yPos)-Step(gHCoverW, gHCoverH), gHcOpenColor, BF
      Else
        Pic(idxP).Line (xPos, yPos)-Step(gHCoverW, gHCoverH), gHcFoldColor, BF
      End If
    Next j
  Next i
  For i = 1 To iMaxRows * 4
    xPos = iStartX + (i - 1) * gHCoverW
    Pic(idxP).Line (xPos, yPos)-Step(gHCoverW, gHCoverH), vbBlack, B
  Next i
  
  
  'Hold 그리기
  Pic(idxP).DrawWidth = 1
  iStartX = gMarginL + (iMaxRows * (gYcW + gCellGapW) - gVessel.iHMaxRows * (gYcW + gCellGapW)) / 2
  iStartY = gMarginT + gBayNoH + (gVessel.iDMaxTiers + 1 + 1) * (gYcH + gCellGapH) + gHCoverH + gHCoverGap * 2
  For i = 1 To gVessel.iHMaxRows
    xPos = iStartX + (i - 1) * (gYcW + gCellGapW)
    For j = 1 To gVessel.iHMaxTiers
      iTierIdx = gVessel.iHMaxTiers - (j - 1)
      yPos = iStartY + (j - 1) * (gYcH + gCellGapH)
      
      If gVessel.gBay(iBayIdx).gCell(0, i, iTierIdx).sSS = "Y" Then
        Pic(idxP).Line (xPos, yPos)-Step(gYcW, gYcH), gBColor, BF
      End If
      
      If gVessel.gBay(iBayIdx).gCell(0, i, iTierIdx).sGuide = "Y" Then
        Pic(idxP).Line (xPos, yPos)-Step(gYcW, gYcH), gCellGuideColor, B
      Else
        Pic(idxP).Line (xPos, yPos)-Step(gYcW, gYcH), vbBlack, B
      End If
      
      gVessel.gBay(iBayIdx).gCell(0, i, iTierIdx).iL = xPos
      gVessel.gBay(iBayIdx).gCell(0, i, iTierIdx).iT = yPos
      gVessel.gBay(iBayIdx).gCell(0, i, iTierIdx).iW = gYcW
      gVessel.gBay(iBayIdx).gCell(0, i, iTierIdx).iH = gYcH
      
    Next j
  Next i
  
  Pic(idxP).Visible = True
End Sub

'베이수 만큼의 픽쳐박스 생성
Private Sub SetLoadControls()
  Dim i%
  
  'Hatch
  For i = 0 To picHatch.Count - 1
    picHatch(i).Visible = False
  Next i

  For i = picHatch.Count To gVessel.iNoHatchs - 1
    Load picHatch(i)
    picHatch(i).AutoRedraw = True
    Set picHatch(i).Container = picVessel
    picHatch(i).Visible = False
  Next i
  
  'Bay
  For i = 0 To Pic.Count - 1
    Pic(i).Visible = False
    
    If i + 1 <= gVessel.iNoBays Then
      Set Pic(i).Container = picHatch(gVessel.gBay(i + 1).iHchNo - 1)
    End If
  Next i
  
  For i = Pic.Count To gVessel.iNoBays - 1
    Load Pic(i)
    Pic(i).AutoRedraw = True
    Set Pic(i).Container = picHatch(gVessel.gBay(i + 1).iHchNo - 1)
    Pic(i).Visible = False
    
    Load shpBox(i)
    Set shpBox(i).Container = Pic(i)
    shpBox(i).Visible = False
    
    Load txtSNo(i)
    Set txtSNo(i).Container = Pic(i)
    txtSNo(i).Visible = False
    
    Load txtNStkWgt(i)
    Set txtNStkWgt(i).Container = Pic(i)
    txtNStkWgt(i).Visible = False
  Next i
End Sub

Private Sub optBay_Click()
  gSynchro = "B"
End Sub

Private Sub optHatch_Click()
  gSynchro = "H"
End Sub

Private Sub Pic_Mousedown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  gStartX = X: gStartY = Y
  
  'Cell/ H Cover Draw Mode
  If Button <> 1 And Button <> 2 Then Exit Sub
  
  shpBox(Index).Left = X: shpBox(Index).Top = Y
  shpBox(Index).Width = 1: shpBox(Index).Height = 1
  shpBox(Index).Visible = True
  
  txtSNo(Index).Visible = False
  txtNStkWgt(Index).Visible = False
End Sub

Private Sub Pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 1 And Button <> 2 Then Exit Sub
  
  If X < gStartX Then shpBox(Index).Left = X
  If Y < gStartY Then shpBox(Index).Top = Y
  shpBox(Index).Width = Abs(X - gStartX): shpBox(Index).Height = Abs(Y - gStartY)
End Sub

Private Sub Pic_Mouseup(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i%, j%, k%, l%
  Dim iBayIdx%, iHatchIdx%
  Dim sXPos1%, sXPos2%, sYPos1%, sYPos2%
  Dim xPos1%, xPos2%, yPos1%, yPos2%
  Dim nHcF As Single, nHcT As Single
  Dim iCvrIdx%, bHCvrFlag As Boolean
  
  If Button <> 1 And Button <> 2 Then Exit Sub
  shpBox(Index).Visible = False
  
  'Cell Info Draw
  iBayIdx = Index + 1
  iHatchIdx = gVessel.gBay(iBayIdx).iHchNo
  
  sXPos1 = shpBox(Index).Left: sXPos2 = shpBox(Index).Left + shpBox(Index).Width
  sYPos1 = shpBox(Index).Top: sYPos2 = shpBox(Index).Top + shpBox(Index).Height
  
  For i = 1 To gVessel.iNoBays
    'Bay Synchro
    If gDrawMode = "B" Or gSynchro = "B" Then
      If gVessel.gBay(i).iBay = iBayIdx Then
      
        For j = 0 To 1
          For k = 1 To gVessel.gBay(i).iNoRows(j)
            For l = 1 To gVessel.gBay(i).iNoTiers(j)
              xPos1 = gVessel.gBay(i).gCell(j, k, l).iL
              xPos2 = gVessel.gBay(i).gCell(j, k, l).iL + gVessel.gBay(i).gCell(j, k, l).iW
              yPos1 = gVessel.gBay(i).gCell(j, k, l).iT
              yPos2 = gVessel.gBay(i).gCell(j, k, l).iT + gVessel.gBay(i).gCell(j, k, l).iH
              
              If Not (xPos1 > sXPos2 Or xPos2 < sXPos1) Then
                If Not (yPos1 > sYPos2 Or yPos2 < sYPos1) Then
                  If Button = 1 Then
                    If gDrawCell = "C" Then
                      gVessel.gBay(i).gCell(j, k, l).sSS = "Y"
                    Else
                      If gVessel.gBay(i).gCell(j, k, l).sSS = "Y" And gVessel.gBay(i).sSize = "2" Then
                        gVessel.gBay(i).gCell(j, k, l).sGuide = "Y"
                      End If
                    End If
                    
                  Else
                    If gDrawCell = "C" Then
                      gVessel.gBay(i).gCell(j, k, l).sSS = ""
                    Else
                      gVessel.gBay(i).gCell(j, k, l).sGuide = ""
                    End If
                    
                  End If
                  
                End If
              End If
                
            Next l
          Next k
        Next j
      End If
      
    'Hatch Synchro
    Else
      If gVessel.gBay(i).iHchNo = iHatchIdx Then
      
        For j = 0 To 1
          For k = 1 To gVessel.gBay(i).iNoRows(j)
            For l = 1 To gVessel.gBay(i).iNoTiers(j)
              xPos1 = gVessel.gBay(i).gCell(j, k, l).iL
              xPos2 = gVessel.gBay(i).gCell(j, k, l).iL + gVessel.gBay(i).gCell(j, k, l).iW
              yPos1 = gVessel.gBay(i).gCell(j, k, l).iT
              yPos2 = gVessel.gBay(i).gCell(j, k, l).iT + gVessel.gBay(i).gCell(j, k, l).iH
              
              If Not (xPos1 > sXPos2 Or xPos2 < sXPos1) Then
                If Not (yPos1 > sYPos2 Or yPos2 < sYPos1) Then
                  If Button = 1 Then
                    If gDrawCell = "C" Then
                      gVessel.gBay(i).gCell(j, k, l).sSS = "Y"
                    Else
                      If gVessel.gBay(i).gCell(j, k, l).sSS = "Y" And gVessel.gBay(i).sSize = "2" Then
                        gVessel.gBay(i).gCell(j, k, l).sGuide = "Y"
                      End If
                    End If
                    
                  Else
                    If gDrawCell = "C" Then
                      gVessel.gBay(i).gCell(j, k, l).sSS = ""
                    Else
                      gVessel.gBay(i).gCell(j, k, l).sGuide = ""
                    End If
                    
                  End If
                  
                End If
              End If
                
            Next l
          Next k
        Next j
      End If
    
    End If
  Next i
  
  '셀복사
  For i = 1 To gVessel.iNoBays
    For j = 0 To 1
      Call PasteCellInfo(i, i, j, "O", gVessel.iMaxRows, gVessel.iMaxTiers)
    Next j
  Next i
  
  
  'Hatch Cover
  yPos1 = gMarginT + gBayNoH + (gVessel.iDMaxTiers + 1 + 1) * (gYcH + gCellGapH)
  yPos2 = yPos1 + gHCoverH + gHCoverGap
  
  iCvrIdx = 0
  For i = 1 To gVessel.iMaxRows * 4
    xPos1 = gMarginL + (i - 1) * gHCoverW
    xPos2 = xPos1 + gHCoverW
    
    If Not (xPos1 > sXPos2 Or xPos2 < sXPos1) Then
      If Not (yPos1 > sYPos2 Or yPos2 < sYPos1) Then
        iCvrIdx = iCvrIdx + 1
        If iCvrIdx = 1 Then
          nHcF = 0.25 * (i - 1)
        End If
        nHcT = 0.25 * i
      End If
    End If
  Next i
  
  'Hatch Cover 그리기
  If Button = 1 And sYPos1 >= yPos1 And sYPos2 <= yPos2 Then
    'Bay Synchro
    If gDrawMode = "B" Or gSynchro = "B" Then
      
      bHCvrFlag = False
      For i = 1 To gVessel.gBay(iBayIdx).iNoCvr
        If Not (gVessel.gBay(iBayIdx).gHchCvr(i).nF > nHcT Or gVessel.gBay(iBayIdx).gHchCvr(i).nT < nHcF) Then
          bHCvrFlag = True
          
          If nHcF < gVessel.gBay(iBayIdx).gHchCvr(i).nF Then
            gVessel.gBay(iBayIdx).gHchCvr(i).nF = nHcF
          End If
          If nHcT > gVessel.gBay(iBayIdx).gHchCvr(i).nT Then
            gVessel.gBay(iBayIdx).gHchCvr(i).nT = nHcT
          End If
          
          Exit For
        End If
      Next i
        
      If bHCvrFlag Then
        For i = 1 To gVessel.gBay(iBayIdx).iNoCvr
          If i > gVessel.gBay(iBayIdx).iNoCvr Then Exit For
          nHcF = gVessel.gBay(iBayIdx).gHchCvr(i).nF
          nHcT = gVessel.gBay(iBayIdx).gHchCvr(i).nT
          
          For j = i + 1 To gVessel.gBay(iBayIdx).iNoCvr
            If Not (gVessel.gBay(iBayIdx).gHchCvr(j).nF > nHcT Or gVessel.gBay(iBayIdx).gHchCvr(j).nT < nHcF) Then
              gVessel.gBay(iBayIdx).gHchCvr(i).nF = nHcF
              gVessel.gBay(iBayIdx).gHchCvr(i).nT = gVessel.gBay(iBayIdx).gHchCvr(j).nT
              
              For k = j + 1 To gVessel.gBay(iBayIdx).iNoCvr
                gVessel.gBay(iBayIdx).gHchCvr(k - 1).iCvr = gVessel.gBay(iBayIdx).gHchCvr(k).iCvr
                gVessel.gBay(iBayIdx).gHchCvr(k - 1).iBay = gVessel.gBay(iBayIdx).gHchCvr(k).iBay
                gVessel.gBay(iBayIdx).gHchCvr(k - 1).nF = gVessel.gBay(iBayIdx).gHchCvr(k).nF
                gVessel.gBay(iBayIdx).gHchCvr(k - 1).nT = gVessel.gBay(iBayIdx).gHchCvr(k).nT
                gVessel.gBay(iBayIdx).gHchCvr(k - 1).sHType = gVessel.gBay(iBayIdx).gHchCvr(k).sHType
              Next k
            
              gVessel.gBay(iBayIdx).iNoCvr = gVessel.gBay(iBayIdx).iNoCvr - 1
              ReDim Preserve gVessel.gBay(iBayIdx).gHchCvr(1 To gVessel.gBay(iBayIdx).iNoCvr)
              Exit For
            End If
            
          Next j
        Next i
        
      Else
        gVessel.gBay(iBayIdx).iNoCvr = gVessel.gBay(iBayIdx).iNoCvr + 1
        ReDim Preserve gVessel.gBay(iBayIdx).gHchCvr(1 To gVessel.gBay(iBayIdx).iNoCvr)
        
        For i = gVessel.gBay(iBayIdx).iNoCvr To 1 Step -1
          If gVessel.gBay(iBayIdx).gHchCvr(i).nF <= nHcF Then
            For j = gVessel.gBay(iBayIdx).iNoCvr - 1 To i Step -1
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).iCvr = gVessel.gBay(iBayIdx).gHchCvr(j).iCvr
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).iBay = gVessel.gBay(iBayIdx).gHchCvr(j).iBay
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).nF = gVessel.gBay(iBayIdx).gHchCvr(j).nF
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).nT = gVessel.gBay(iBayIdx).gHchCvr(j).nT
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).sHType = gVessel.gBay(iBayIdx).gHchCvr(j).sHType
            Next j
            
            gVessel.gBay(iBayIdx).gHchCvr(i).iCvr = i
            gVessel.gBay(iBayIdx).gHchCvr(i).iBay = iBayIdx
            gVessel.gBay(iBayIdx).gHchCvr(i).nF = nHcF
            gVessel.gBay(iBayIdx).gHchCvr(i).nT = nHcT
            gVessel.gBay(iBayIdx).gHchCvr(i).sHType = gDrawHCover
            
            Exit For
          End If
        Next i
      End If
        
    'Hatch Synchro
    Else
      For l = 1 To gVessel.iNoBays
        If gVessel.gBay(l).iHchNo = iHatchIdx Then
          iCvrIdx = 0
          For i = 1 To gVessel.iMaxRows * 4
            xPos1 = gMarginL + (i - 1) * gHCoverW
            xPos2 = xPos1 + gHCoverW
            
            If Not (xPos1 > sXPos2 Or xPos2 < sXPos1) Then
              If Not (yPos1 > sYPos2 Or yPos2 < sYPos1) Then
                iCvrIdx = iCvrIdx + 1
                If iCvrIdx = 1 Then
                  nHcF = 0.25 * (i - 1)
                End If
                nHcT = 0.25 * i
              End If
            End If
          Next i
          
          bHCvrFlag = False
          For i = 1 To gVessel.gBay(l).iNoCvr
            If Not (gVessel.gBay(l).gHchCvr(i).nF > nHcT Or gVessel.gBay(l).gHchCvr(i).nT < nHcF) Then
              bHCvrFlag = True
              
              If nHcF < gVessel.gBay(l).gHchCvr(i).nF Then
                gVessel.gBay(l).gHchCvr(i).nF = nHcF
              End If
              If nHcT > gVessel.gBay(l).gHchCvr(i).nT Then
                gVessel.gBay(l).gHchCvr(i).nT = nHcT
              End If
              
              Exit For
            End If
          Next i
          
          If bHCvrFlag Then
            For i = 1 To gVessel.gBay(l).iNoCvr
              If i > gVessel.gBay(l).iNoCvr Then Exit For
              nHcF = gVessel.gBay(l).gHchCvr(i).nF
              nHcT = gVessel.gBay(l).gHchCvr(i).nT
              
              For j = i + 1 To gVessel.gBay(l).iNoCvr
                If Not (gVessel.gBay(l).gHchCvr(j).nF > nHcT Or gVessel.gBay(l).gHchCvr(j).nT < nHcF) Then
                  gVessel.gBay(l).gHchCvr(i).nF = nHcF
                  gVessel.gBay(l).gHchCvr(i).nT = gVessel.gBay(l).gHchCvr(j).nT
                  
                  For k = j + 1 To gVessel.gBay(l).iNoCvr
                    gVessel.gBay(l).gHchCvr(k - 1).iCvr = gVessel.gBay(l).gHchCvr(k).iCvr
                    gVessel.gBay(l).gHchCvr(k - 1).iBay = gVessel.gBay(l).gHchCvr(k).iBay
                    gVessel.gBay(l).gHchCvr(k - 1).nF = gVessel.gBay(l).gHchCvr(k).nF
                    gVessel.gBay(l).gHchCvr(k - 1).nT = gVessel.gBay(l).gHchCvr(k).nT
                    gVessel.gBay(l).gHchCvr(k - 1).sHType = gVessel.gBay(l).gHchCvr(k).sHType
                  Next k
                  
                  gVessel.gBay(l).iNoCvr = gVessel.gBay(l).iNoCvr - 1
                  ReDim Preserve gVessel.gBay(l).gHchCvr(1 To gVessel.gBay(l).iNoCvr)
                  Exit For
                End If
              Next j
            Next i
              
          Else
            gVessel.gBay(l).iNoCvr = gVessel.gBay(l).iNoCvr + 1
            ReDim Preserve gVessel.gBay(l).gHchCvr(1 To gVessel.gBay(l).iNoCvr)
            
            For i = gVessel.gBay(l).iNoCvr To 1 Step -1
              If gVessel.gBay(l).gHchCvr(i).nF <= nHcF Then
                For j = gVessel.gBay(l).iNoCvr - 1 To i Step -1
                  gVessel.gBay(l).gHchCvr(j + 1).iCvr = gVessel.gBay(l).gHchCvr(j).iCvr
                  gVessel.gBay(l).gHchCvr(j + 1).iBay = gVessel.gBay(l).gHchCvr(j).iBay
                  gVessel.gBay(l).gHchCvr(j + 1).nF = gVessel.gBay(l).gHchCvr(j).nF
                  gVessel.gBay(l).gHchCvr(j + 1).nT = gVessel.gBay(l).gHchCvr(j).nT
                  gVessel.gBay(l).gHchCvr(j + 1).sHType = gVessel.gBay(l).gHchCvr(j).sHType
                Next j
                
                gVessel.gBay(l).gHchCvr(i).iCvr = i
                gVessel.gBay(l).gHchCvr(i).iBay = l
                gVessel.gBay(l).gHchCvr(i).nF = nHcF
                gVessel.gBay(l).gHchCvr(i).nT = nHcT
                gVessel.gBay(l).gHchCvr(i).sHType = gDrawHCover
                
                Exit For
              End If
            Next i
          End If
          
        End If
      Next l
    End If
  End If
  
  'Hatch Cover 지우기
  If Button = 2 And sYPos1 >= yPos1 And sYPos2 <= yPos2 Then
    'Bay Synchro
    If gDrawMode = "B" Or gSynchro = "B" Then
      For i = 1 To gVessel.gBay(iBayIdx).iNoCvr
        If i > gVessel.gBay(iBayIdx).iNoCvr Then Exit For
        If Not (gVessel.gBay(iBayIdx).gHchCvr(i).nF > nHcT Or gVessel.gBay(iBayIdx).gHchCvr(i).nT < nHcF) Then
          '다 지우는 경우
          If nHcF <= gVessel.gBay(iBayIdx).gHchCvr(i).nF And nHcT >= gVessel.gBay(iBayIdx).gHchCvr(i).nT Then
            For j = i + 1 To gVessel.gBay(iBayIdx).iNoCvr
              gVessel.gBay(iBayIdx).gHchCvr(j - 1).iCvr = gVessel.gBay(iBayIdx).gHchCvr(j).iCvr
              gVessel.gBay(iBayIdx).gHchCvr(j - 1).iBay = gVessel.gBay(iBayIdx).gHchCvr(j).iBay
              gVessel.gBay(iBayIdx).gHchCvr(j - 1).nF = gVessel.gBay(iBayIdx).gHchCvr(j).nF
              gVessel.gBay(iBayIdx).gHchCvr(j - 1).nT = gVessel.gBay(iBayIdx).gHchCvr(j).nT
              gVessel.gBay(iBayIdx).gHchCvr(j - 1).sHType = gVessel.gBay(iBayIdx).gHchCvr(j).sHType
            Next j
            
            gVessel.gBay(iBayIdx).iNoCvr = gVessel.gBay(iBayIdx).iNoCvr - 1
            If gVessel.gBay(iBayIdx).iNoCvr = 0 Then
              Erase gVessel.gBay(iBayIdx).gHchCvr()
              Exit For
            Else
              ReDim Preserve gVessel.gBay(iBayIdx).gHchCvr(1 To gVessel.gBay(iBayIdx).iNoCvr)
            End If
            i = i - 1
        
          '2개로 분리되는 경우
          ElseIf nHcF > gVessel.gBay(iBayIdx).gHchCvr(i).nF And nHcT < gVessel.gBay(iBayIdx).gHchCvr(i).nT Then
            gVessel.gBay(iBayIdx).iNoCvr = gVessel.gBay(iBayIdx).iNoCvr + 1
            ReDim Preserve gVessel.gBay(iBayIdx).gHchCvr(1 To gVessel.gBay(iBayIdx).iNoCvr)
            
            For j = gVessel.gBay(iBayIdx).iNoCvr - 1 To i + 1 Step -1
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).iCvr = gVessel.gBay(iBayIdx).gHchCvr(j).iCvr
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).iBay = gVessel.gBay(iBayIdx).gHchCvr(j).iBay
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).nF = gVessel.gBay(iBayIdx).gHchCvr(j).nF
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).nT = gVessel.gBay(iBayIdx).gHchCvr(j).nT
              gVessel.gBay(iBayIdx).gHchCvr(j + 1).sHType = gVessel.gBay(iBayIdx).gHchCvr(j).sHType
            Next j
        
            gVessel.gBay(iBayIdx).gHchCvr(i + 1).iCvr = i + 1
            gVessel.gBay(iBayIdx).gHchCvr(i + 1).iBay = iBayIdx
            gVessel.gBay(iBayIdx).gHchCvr(i + 1).nF = nHcT
            gVessel.gBay(iBayIdx).gHchCvr(i + 1).nT = gVessel.gBay(iBayIdx).gHchCvr(i).nT
            gVessel.gBay(iBayIdx).gHchCvr(i + 1).sHType = gVessel.gBay(iBayIdx).gHchCvr(i).sHType
        
            gVessel.gBay(iBayIdx).gHchCvr(i).nT = nHcF
        
            i = i + 1
        
          '짧아지는 경우
          Else
            If nHcF <= gVessel.gBay(iBayIdx).gHchCvr(i).nF Then
              gVessel.gBay(iBayIdx).gHchCvr(i).nF = nHcT
            End If
            If nHcT >= gVessel.gBay(iBayIdx).gHchCvr(i).nT Then
              gVessel.gBay(iBayIdx).gHchCvr(i).nT = nHcF
            End If
          End If
        End If
      Next i
      
    'Hatch Synchro
    Else
      For l = 1 To gVessel.iNoBays
        If gVessel.gBay(l).iHchNo = iHatchIdx Then
          For i = 1 To gVessel.gBay(l).iNoCvr
            If i > gVessel.gBay(l).iNoCvr Then Exit For
            If Not (gVessel.gBay(l).gHchCvr(i).nF > nHcT Or gVessel.gBay(l).gHchCvr(i).nT < nHcF) Then
              '다 지우는 경우
              If nHcF <= gVessel.gBay(l).gHchCvr(i).nF And nHcT >= gVessel.gBay(l).gHchCvr(i).nT Then
                For j = i + 1 To gVessel.gBay(l).iNoCvr
                  gVessel.gBay(l).gHchCvr(j - 1).iCvr = gVessel.gBay(l).gHchCvr(j).iCvr
                  gVessel.gBay(l).gHchCvr(j - 1).iBay = gVessel.gBay(l).gHchCvr(j).iBay
                  gVessel.gBay(l).gHchCvr(j - 1).nF = gVessel.gBay(l).gHchCvr(j).nF
                  gVessel.gBay(l).gHchCvr(j - 1).nT = gVessel.gBay(l).gHchCvr(j).nT
                  gVessel.gBay(l).gHchCvr(j - 1).sHType = gVessel.gBay(l).gHchCvr(j).sHType
                Next j
                
                gVessel.gBay(l).iNoCvr = gVessel.gBay(l).iNoCvr - 1
                If gVessel.gBay(l).iNoCvr = 0 Then
                  Erase gVessel.gBay(l).gHchCvr()
                  Exit For
                Else
                  ReDim Preserve gVessel.gBay(l).gHchCvr(1 To gVessel.gBay(l).iNoCvr)
                End If
                i = i - 1
            
              '2개로 분리되는 경우
              ElseIf nHcF > gVessel.gBay(l).gHchCvr(i).nF And nHcT < gVessel.gBay(l).gHchCvr(i).nT Then
                gVessel.gBay(l).iNoCvr = gVessel.gBay(l).iNoCvr + 1
                ReDim Preserve gVessel.gBay(l).gHchCvr(1 To gVessel.gBay(l).iNoCvr)
                
                For j = gVessel.gBay(l).iNoCvr - 1 To i + 1 Step -1
                  gVessel.gBay(l).gHchCvr(j + 1).iCvr = gVessel.gBay(l).gHchCvr(j).iCvr
                  gVessel.gBay(l).gHchCvr(j + 1).iBay = gVessel.gBay(l).gHchCvr(j).iBay
                  gVessel.gBay(l).gHchCvr(j + 1).nF = gVessel.gBay(l).gHchCvr(j).nF
                  gVessel.gBay(l).gHchCvr(j + 1).nT = gVessel.gBay(l).gHchCvr(j).nT
                  gVessel.gBay(l).gHchCvr(j + 1).sHType = gVessel.gBay(l).gHchCvr(j).sHType
                Next j
            
                gVessel.gBay(l).gHchCvr(i + 1).iCvr = i + 1
                gVessel.gBay(l).gHchCvr(i + 1).iBay = l
                gVessel.gBay(l).gHchCvr(i + 1).nF = nHcT
                gVessel.gBay(l).gHchCvr(i + 1).nT = gVessel.gBay(l).gHchCvr(i).nT
                gVessel.gBay(l).gHchCvr(i + 1).sHType = gVessel.gBay(l).gHchCvr(i).sHType
            
                gVessel.gBay(l).gHchCvr(i).nT = nHcF
            
                i = i + 1
            
              '짧아지는 경우
              Else
                If nHcF <= gVessel.gBay(l).gHchCvr(i).nF Then
                  gVessel.gBay(l).gHchCvr(i).nF = nHcT
                End If
                If nHcT >= gVessel.gBay(l).gHchCvr(i).nT Then
                  gVessel.gBay(l).gHchCvr(i).nT = nHcF
                End If
              End If
            End If
          Next i
        End If
      Next l
    End If
  End If
  
  'Row Stacking Weight
  For i = 0 To 1
    For j = 1 To gVessel.gBay(iBayIdx).iNoRows(i)
      xPos1 = gVessel.gBay(iBayIdx).gRow(i, j).iL_Wgt
      xPos2 = xPos1 + gVessel.gBay(iBayIdx).gRow(i, j).iW_Wgt
      yPos1 = gVessel.gBay(iBayIdx).gRow(i, j).iT_Wgt
      yPos2 = yPos1 + gVessel.gBay(iBayIdx).gRow(i, j).iH_Wgt
      
      If gStartX >= xPos1 And gStartX <= xPos2 And X >= xPos1 And X <= xPos2 Then
        If gStartY >= yPos1 And gStartY <= yPos2 And Y >= yPos1 And Y <= yPos2 Then
          
          txtNStkWgt(Index).FontSize = gWgtFontSize
          txtNStkWgt(Index).Move xPos1 - 1, yPos1 - 1, gVessel.gBay(iBayIdx).gRow(i, j).iW_Wgt + 2, gVessel.gBay(iBayIdx).gRow(i, j).iH_Wgt + 2
          txtNStkWgt(Index).Text = Trim(Str(gVessel.gBay(iBayIdx).gRow(i, j).nStkWgt))
          txtNStkWgt(Index).SelLength = Len(txtNStkWgt(Index).Text)
          txtNStkWgt(Index).Visible = True
          txtNStkWgt(Index).SetFocus
          
          gHdIdx = i: gNStkWgtIdx = j
        End If
      End If
    Next j
  Next i
  
  'Row / Tier No.
  For i = 0 To 1
    For j = 1 To gVessel.gBay(iBayIdx).iNoRows(i)
      xPos1 = gVessel.gBay(iBayIdx).gRow(i, j).iL
      xPos2 = xPos1 + gVessel.gBay(iBayIdx).gRow(i, j).iW
      yPos1 = gVessel.gBay(iBayIdx).gRow(i, j).iT
      yPos2 = yPos1 + gVessel.gBay(iBayIdx).gRow(i, j).iH
      
      If gStartX >= xPos1 And gStartX <= xPos2 And X >= xPos1 And X <= xPos2 Then
        If gStartY >= yPos1 And gStartY <= yPos2 And Y >= yPos1 And Y <= yPos2 Then
          
          txtSNo(Index).FontSize = gNoFontSize
          txtSNo(Index).Move xPos1 - 1, yPos1 - 1, gVessel.gBay(iBayIdx).gRow(i, j).iW + 2, gVessel.gBay(iBayIdx).gRow(i, j).iH + 2
          txtSNo(Index).Text = gVessel.gBay(iBayIdx).gRow(i, j).sNo
          txtSNo(Index).SelLength = Len(txtSNo(Index).Text)
          txtSNo(Index).Visible = True
          txtSNo(Index).SetFocus
          
          gRowTier = "R"
          gHdIdx = i: gSnoIdx = j
        End If
      End If
    Next j
    
    For j = 1 To gVessel.gBay(iBayIdx).iNoTiers(i)
      xPos1 = gVessel.gBay(iBayIdx).gTier(i, j).iL
      xPos2 = xPos1 + gVessel.gBay(iBayIdx).gTier(i, j).iW
      yPos1 = gVessel.gBay(iBayIdx).gTier(i, j).iT
      yPos2 = yPos1 + gVessel.gBay(iBayIdx).gTier(i, j).iH
      
      If gStartX >= xPos1 And gStartX <= xPos2 And X >= xPos1 And X <= xPos2 Then
        If gStartY >= yPos1 And gStartY <= yPos2 And Y >= yPos1 And Y <= yPos2 Then
          
          txtSNo(Index).FontSize = gNoFontSize
          txtSNo(Index).Move xPos1 - 1, yPos1 - 1, gVessel.gBay(iBayIdx).gTier(i, j).iW + 2, gVessel.gBay(iBayIdx).gTier(i, j).iH + 2
          txtSNo(Index).Text = gVessel.gBay(iBayIdx).gTier(i, j).sNo
          txtSNo(Index).SelLength = Len(txtSNo(Index).Text)
          txtSNo(Index).Visible = True
          txtSNo(Index).SetFocus
          
          gRowTier = "T"
          gHdIdx = i: gSnoIdx = j
        End If
      End If
    Next j
  Next i
  
  gHatchIdx = iHatchIdx: gBayIdx = iBayIdx
  
  '다시 그리기
  If txtSNo(Index).Visible = False And txtNStkWgt(Index).Visible = False Then Call ReDraw
End Sub

Private Sub ReDraw()
  Screen.MousePointer = 11
  If gDrawMode = "A" Then
    Call SetDrawStructure
  ElseIf gDrawMode = "H" Then
    Call SetDrawHatch(gDrawNo)
  ElseIf gDrawMode = "B" Then
    Call SetDrawBay(gDrawNo)
  End If
  Screen.MousePointer = 0
End Sub

Private Sub sprList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  If Col = 2 And Row > 0 Then
    sprList.Row = Row: sprList.Col = Col
    
    If Len(Trim(sprList.Text)) = 2 Then
      If Val(sprList.Text) Mod 2 = 0 Then
        sprList.SetText 3, Row, "40"
      Else
        sprList.SetText 3, Row, "20"
      End If
    End If
  End If
End Sub

Private Sub tbrBay_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim iRow%
  Dim i%, j%
  
  iRow = sprList.SelBlockRow
  If iRow <= 0 Then
    Exit Sub
  End If
  
  Select Case Button.Key
    Case "imgInsert"
      sprList.MaxRows = sprList.MaxRows + 1
      
      sprList.Row = sprList.MaxRows
      For j = 1 To sprList.MaxCols
        sprList.Col = j
        sprList.CellType = CellTypeEdit
        sprList.TypeHAlign = TypeHAlignCenter
        sprList.TypeVAlign = TypeVAlignCenter
        sprList.TypeMaxEditLen = 2
        sprList.TypeEditCharSet = TypeEditCharSetNumeric
        sprList.EditEnterAction = EditEnterActionDown
        sprList.EditModeReplace = True
      Next j
      
      For i = sprList.MaxRows - 1 To iRow + 1 Step -1
        sprList.Row = i
        For j = 1 To sprList.MaxCols
          sprList.Col = j
          sprList.SetText j, i + 1, sprList.Text
        Next j
      Next i
      
      For j = 1 To sprList.MaxCols
        sprList.SetText j, iRow + 1, ""
      Next j
      
    Case "imgRemove"
      For i = iRow + 1 To sprList.MaxRows
        sprList.Row = i
        For j = 1 To sprList.MaxCols
          sprList.Col = j
          sprList.SetText j, i - 1, sprList.Text
        Next j
      Next i
      
      sprList.MaxRows = sprList.MaxRows - 1
  End Select
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "imgRefresh"
      Call GetVslStructure(gVessel.sVCode)
      Call SetTreeViewItem
      Call ReDraw
      tbrMain.Buttons(32).Enabled = False
      
      Call SetVslInfo
      Call SetBayInfo
      
    Case "imgHide"
      Call HideTreeView
    Case "imgExpand"
      Call ExpandTreeView
      
    Case "imgClose"
      Unload Me
      
    Case "imgZoomOut"
      If gZoomIdx > 0.5 Then
        gZoomIdx = gZoomIdx - 0.25
        txtZoom.Text = Trim(Str(gZoomIdx * 100)) & "%"
        
        Call ReDraw
        Call ReSizePic
      End If
    Case "imgZoomIn"
      If gZoomIdx < 2 Then
        gZoomIdx = gZoomIdx + 0.25
        txtZoom.Text = Trim(Str(gZoomIdx * 100)) & "%"
        
        Call ReDraw
        Call ReSizePic
      End If
    Case "imgZoom"
      gZoomIdx = 1
      txtZoom.Text = Trim(Str(gZoomIdx * 100)) & "%"
      
      Call ReDraw
      Call ReSizePic
      
    Case "imgDraw"
      gDrawCell = "C"
    Case "ImgCellGuide"
      gDrawCell = "G"
    Case "imgHOpen"
      gDrawHCover = "O"
    Case "imgHFold"
      gDrawHCover = "F"
    
    Case "imgCopy"
      Call CopyStructure
      
    Case "imgPaste"
      Call PasteStructure
      
  End Select
End Sub

Private Sub CopyStructure()
  'Copy
  gCopyHatchIdx = gHatchIdx: gCopyBayIdx = gBayIdx
  
  If gCopyHatchIdx > 0 Then
    tbrMain.Buttons(32).Enabled = True
  Else
    tbrMain.Buttons(32).Enabled = False
  End If
End Sub

Private Sub PasteStructure()
  Dim i%
  Dim iNoCopyBay%, iNoPasteBay%
  Dim iCopyBay() As Integer, iPasteBay() As Integer
  
  'Paste
  Select Case gCopyMode
    Case "H"
      For i = 1 To gVessel.iNoBays
        If gVessel.gBay(i).iHchNo = gCopyHatchIdx Then
          iNoCopyBay = iNoCopyBay + 1
          ReDim Preserve iCopyBay(1 To iNoCopyBay)
          iCopyBay(iNoCopyBay) = gVessel.gBay(i).iBay
        End If
        
        If gVessel.gBay(i).iHchNo = gHatchIdx Then
          iNoPasteBay = iNoPasteBay + 1
          ReDim Preserve iPasteBay(1 To iNoPasteBay)
          iPasteBay(iNoPasteBay) = gVessel.gBay(i).iBay
        End If
      Next i
      
      If iNoPasteBay <> iNoCopyBay Then
        MsgBox "Can't Copy!! Bay Count is different from copied bays in Hatch."
        Exit Sub
      End If
      
      For i = 1 To iNoCopyBay
        Call PasteCover(iCopyBay(i), iPasteBay(i))
        
        Call PasteCellInfo(iCopyBay(i), iPasteBay(i), 1)
        Call PasteCellInfo(iCopyBay(i), iPasteBay(i), 0)
        
        Call PasteNo(iCopyBay(i), iPasteBay(i), 1)
        Call PasteNo(iCopyBay(i), iPasteBay(i), 0)
        
      Next i
    
    Case "HC"
     Call PasteCover(gCopyBayIdx, gBayIdx)
      
    Case "B"
      Call PasteCellInfo(gCopyBayIdx, gBayIdx, 1)
      Call PasteCellInfo(gCopyBayIdx, gBayIdx, 0)
      
      Call PasteNo(gCopyBayIdx, gBayIdx, 1)
      Call PasteNo(gCopyBayIdx, gBayIdx, 0)
      
    Case "D"
      Call PasteCellInfo(gCopyBayIdx, gBayIdx, 1)
      
    Case "HD"
      Call PasteCellInfo(gCopyBayIdx, gBayIdx, 0)
      
    Case "DN"
      Call PasteNo(gCopyBayIdx, gBayIdx, 1)
      
    Case "HN"
      Call PasteNo(gCopyBayIdx, gBayIdx, 0)
      
  End Select
  
  Call ReDraw
End Sub

Private Sub ExpandTreeView()
  Dim i%
  
  If bFlagExpand Then
    bFlagExpand = False
    tbrMain.Buttons(3).ToolTipText = "Expand Vessel Structure List"
    
    For i = 1 To tvList.Nodes.Count
      tvList.Nodes(i).Expanded = False
    Next i
  Else
    bFlagExpand = True
    tbrMain.Buttons(3).ToolTipText = "Collapse Vessel Structure List"
    
    For i = 1 To tvList.Nodes.Count
      tvList.Nodes(i).Expanded = True
    Next i
  End If
End Sub

Private Sub HideTreeView()
  If bFlagHide Then
    bFlagHide = False
    tbrMain.Buttons(2).ToolTipText = "Hide Vessel Structure List"
    
    picBack.Move picBackCenter.Left + picBackCenter.Width, picBack.Top, Me.ScaleWidth - tvList.Width - picBackCenter.Width - 2, picBack.Height
  Else
    bFlagHide = True
    tbrMain.Buttons(2).ToolTipText = "Show Vessel Structure List"
    
    picBack.Move 0, picBack.Top, Me.ScaleWidth - 2, picBack.Height
  End If
  
  fraTable.Move 0, 0, picBack.Width, picBack.Height
  fraBasic.Move 0, 0, picBack.Width, picBack.Height
End Sub

Private Sub ReSizePic()
  picVessel.Move 0, 0, gPicW + vsVsl.Width, gPicH + hsVsl.Height
  Call SetScroll
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Select Case ButtonMenu.Key
    Case "chkHatch"
      txtCopy.Text = "Copy Hatch"
      gCopyMode = "H"
      tbrMain.Buttons(32).Enabled = False
    Case "chkBay"
      txtCopy.Text = "Copy Bay"
      gCopyMode = "B"
      tbrMain.Buttons(32).Enabled = False
    Case "chkCover"
      txtCopy.Text = "Copy Cover"
      gCopyMode = "HC"
      tbrMain.Buttons(32).Enabled = False
    Case "chkDeck"
      txtCopy.Text = "Copy Deck"
      gCopyMode = "D"
      tbrMain.Buttons(32).Enabled = False
    Case "chkHold"
      txtCopy.Text = "Copy Hold"
      gCopyMode = "HD"
      tbrMain.Buttons(32).Enabled = False
    Case "chkDeckNo"
      txtCopy.Text = "Copy Deck No."
      gCopyMode = "DN"
      tbrMain.Buttons(32).Enabled = False
    Case "chkHoldNo"
      txtCopy.Text = "Copy Hold No."
      gCopyMode = "HN"
      tbrMain.Buttons(32).Enabled = False
  End Select
End Sub

Private Function ChkBasicValidation() As Boolean
  Dim i%, j%, iTmpH%, iTmpB%
  Dim iHch%, iBay%, sSz$
  Dim iNoBays%, iNoHchs%
  Dim iVal%, iVal2%
  
  ChkBasicValidation = True
  
  '유효한지 확인 - 해치 & 베이 & 사이즈 체크
  For i = 1 To sprList.MaxRows
    sprList.Row = i
    
    '해치 seq 체크
    sprList.Col = 1
    iHch = Val(sprList.Text)
    
    If iHch > 0 And iHch <= gVessel.iNoHatchs And iHch >= iTmpH Then
      If iHch <> iTmpH Then
        '해치 내 베이 수가 1개 ~ 3개 가 아니면 에러
        If iTmpH <> 0 Then
          If iNoBays < 1 And iNoBays > 3 Then
            ChkBasicValidation = False
            Exit Function
          End If
        End If
        
        iNoBays = 1
        iTmpH = iHch
        
        iNoHchs = iNoHchs + 1
      Else
        iNoBays = iNoBays + 1
      End If
    Else
      ChkBasicValidation = False
      Exit Function
    End If
    
    '베이 seq 체크 - 99일 땐 제외
    sprList.Col = 2
    iBay = Val(sprList.Text)
    
    If iBay > 0 And iBay > iTmpB Then
      iTmpB = iBay
    Else
      If iTmpB <> 99 Then
        ChkBasicValidation = False
        Exit Function
      Else
        iTmpB = iBay
      End If
    End If
    
    'size 체크
    sprList.Col = 3
    sSz = Trim(sprList.Text)
    If sSz <> "20" And sSz <> "40" Then
      ChkBasicValidation = False
      Exit Function
    End If
    
    'bay no chk
    If sSz = "20" And iBay Mod 2 = 0 Then
      ChkBasicValidation = False
      Exit Function
    End If
    
    If sSz = "40" And iBay Mod 2 = 1 Then
      ChkBasicValidation = False
      Exit Function
    End If
  Next i
  
  If iNoBays < 1 And iNoBays > 3 Then
    ChkBasicValidation = False
    Exit Function
  End If
  
  '전체 해치개수만큼 해치 정보가 존재
  If iNoHchs <> gVessel.iNoHatchs Then
    ChkBasicValidation = False
    Exit Function
  End If
  
  
  '변경 값 반영
  If ChkBasicValidation Then
    iNoBays = gVessel.iNoBays
    gVessel.iNoBays = sprList.MaxRows
    
    ReDim Preserve gVessel.gBay(1 To gVessel.iNoBays)
    
    For i = 1 To gVessel.iNoBays
      sprList.Row = i
      sprList.Col = 1: iHch = Val(sprList.Text)
      sprList.Col = 2: iBay = Val(sprList.Text)
      sprList.Col = 3: sSz = Left(sprList.Text, 1)
      
      gVessel.gBay(i).iBay = i
      gVessel.gBay(i).iHchNo = iHch
      gVessel.gBay(i).sBayNo = Format(iBay, "00")
      gVessel.gBay(i).sSize = sSz
      
      If i > iNoBays Then
        gVessel.gBay(i).iNoRows(0) = gVessel.iHMaxRows
        gVessel.gBay(i).iNoRows(1) = gVessel.iDMaxRows
        gVessel.gBay(i).iNoTiers(0) = gVessel.iHMaxTiers
        gVessel.gBay(i).iNoTiers(1) = gVessel.iDMaxTiers
        
        'row info
        ReDim gVessel.gBay(i).gRow(0 To 1, 1 To gVessel.iMaxRows)
        For j = 1 To gVessel.iHMaxRows
          gVessel.gBay(i).gRow(0, j).iHD = 0
          gVessel.gBay(i).gRow(0, j).iRow = j
          
          iVal = Int(gVessel.iHMaxRows / 2)
          If gVessel.iHMaxRows Mod 2 = 0 Then
            If j <= iVal Then
              iVal2 = gVessel.iHMaxRows - (j - 1) * 2
              gVessel.gBay(i).gRow(0, j).sNo = Format(iVal2, "00")
            Else
              iVal2 = (j - iVal) * 2 - 1
              gVessel.gBay(i).gRow(0, j).sNo = Format(iVal2, "00")
            End If
            
          Else
            If j <= iVal Then
              iVal2 = (gVessel.iHMaxRows - 1) - (j - 1) * 2
              gVessel.gBay(i).gRow(0, j).sNo = Format(iVal2, "00")
            ElseIf j = iVal + 1 Then
              gVessel.gBay(i).gRow(0, j).sNo = "00"
            Else
              iVal2 = (j - iVal - 1) * 2 - 1
              gVessel.gBay(i).gRow(0, j).sNo = Format(iVal2, "00")
            End If
            
          End If
        Next j
        
        For j = 1 To gVessel.iDMaxRows
          gVessel.gBay(i).gRow(1, j).iHD = 1
          gVessel.gBay(i).gRow(1, j).iRow = j
          
          iVal = Int(gVessel.iDMaxRows / 2)
          If gVessel.iDMaxRows Mod 2 = 0 Then
            If j <= iVal Then
              iVal2 = gVessel.iDMaxRows - (j - 1) * 2
              gVessel.gBay(i).gRow(1, j).sNo = Format(iVal2, "00")
            Else
              iVal2 = (j - iVal) * 2 - 1
              gVessel.gBay(i).gRow(1, j).sNo = Format(iVal2, "00")
            End If
            
          Else
            If j <= iVal Then
              iVal2 = (gVessel.iDMaxRows - 1) - (j - 1) * 2
              gVessel.gBay(i).gRow(1, j).sNo = Format(iVal2, "00")
            ElseIf j = iVal + 1 Then
              gVessel.gBay(i).gRow(1, j).sNo = "00"
            Else
              iVal2 = (j - iVal - 1) * 2 - 1
              gVessel.gBay(i).gRow(1, j).sNo = Format(iVal2, "00")
            End If
            
          End If
        Next j
        
        'Tier Info
        ReDim gVessel.gBay(i).gTier(0 To 1, 1 To gVessel.iMaxTiers)
        For j = 1 To gVessel.iHMaxTiers
          gVessel.gBay(i).gTier(0, j).iHD = 0
          gVessel.gBay(i).gTier(0, j).iTier = j
          iVal = j * 2
          gVessel.gBay(i).gTier(0, j).sNo = Format(iVal, "00")
        Next j
        
        For j = 1 To gVessel.iDMaxTiers
          gVessel.gBay(i).gTier(1, j).iHD = 1
          gVessel.gBay(i).gTier(1, j).iTier = j
          iVal = j * 2 + 80
          gVessel.gBay(i).gTier(1, j).sNo = Format(iVal, "00")
        Next j
        
        'Cell Info
        ReDim gVessel.gBay(i).gCell(0 To 1, 1 To gVessel.iMaxRows, 1 To gVessel.iMaxTiers)
        ReDim gVessel.gBay(i).gCellOrg(0 To 1, 1 To gVessel.iMaxRows, 1 To gVessel.iMaxTiers)
        
      End If
      
    Next i
    
    '트리 초기화
  Call SetTreeViewItem
  End If
  
End Function

Private Sub tvList_KeyUp(KeyCode As Integer, Shift As Integer)
  If Shift <> 2 Then Exit Sub
  
  Select Case KeyCode
    Case vbKeyC
      Call CopyStructure
        
    Case vbKeyV
      If tbrMain.Buttons(32).Enabled Then
        Call PasteStructure
      End If
      
    Case vbKeyZ
      
  End Select
End Sub

Private Sub tvList_NodeClick(ByVal Node As MSComctlLib.Node)
  Dim i%
  
  If fraBasic.Visible Then
    If ChkBasicValidation = False Then
      MsgBox "Warning. Invalid Bay & Hatch Information exist!"
      Exit Sub
    End If
  End If
  
  Call SetLoadControls
  
  picVessel.Cls
  gDrawMode = ""
  imgLeft.Visible = False: imgCenter.Visible = False: imgRight.Visible = False: imgBridge.Visible = False
  fraTable.Visible = False: fraBasic.Visible = False
  If Node.Key = gVessel.sVCode Then
    'Vessel View
    gDrawMode = "V"
    imgLeft.Visible = True: imgCenter.Visible = True: imgRight.Visible = True: imgBridge.Visible = True
    Call SetDrawVessel
    
  ElseIf Node.Key = gVessel.sVCode & "S" Then
    gDrawMode = "A"
    Call SetDrawStructure
    
  ElseIf Right(Node.Text, 5) = "Hatch" Then
    gDrawMode = "H"
    gDrawNo = Right(Node.Key, 2)
    gHatchIdx = Val(gDrawNo)
    
    For i = 1 To gVessel.iNoBays
      If gHatchIdx = gVessel.gBay(i).iHchNo Then
        gBayIdx = gVessel.gBay(i).iBay
        Exit For
      End If
    Next i
    
    Call SetDrawHatch(gDrawNo)
    
  ElseIf Right(Node.Text, 3) = "Bay" Then
    gDrawMode = "B"
    gDrawNo = Right(Node.Key, 2)
    gHatchIdx = GetHatchIdx(gDrawNo)
    gBayIdx = GetBayIdx(gDrawNo)
    
    Call SetDrawBay(gDrawNo)
    
  ElseIf Node.Key = gVessel.sVCode & "T" Then
    Call SetVslInfo
    fraTable.Visible = True
    txtVslCd.SetFocus
  ElseIf Node.Key = gVessel.sVCode & "TB" Then
    Call SetBayInfo
    fraBasic.Visible = True
    
  End If
  
  Call ReSizePic
End Sub

Private Sub InitSpread()
  Dim i%, j%
  
  sprList.SetText 1, 0, "Hatch No."
  sprList.SetText 2, 0, "Bay No."
  sprList.SetText 3, 0, "Size"
  
  sprList.MaxRows = gVessel.iNoBays
  
  For i = 1 To sprList.MaxRows
    sprList.Row = i
    For j = 1 To sprList.MaxCols
      sprList.Col = j
      sprList.CellType = CellTypeEdit
      sprList.TypeHAlign = TypeHAlignCenter
      sprList.TypeVAlign = TypeVAlignCenter
      sprList.TypeMaxEditLen = 2
      sprList.TypeEditCharSet = TypeEditCharSetNumeric
      sprList.EditEnterAction = EditEnterActionDown
      sprList.EditModeReplace = True
      
      sprList.SetText j, i, ""
    Next j
  Next i
  
End Sub

Private Sub SetBayInfo()
  Dim i%
  
  Call InitSpread
  
  For i = 1 To gVessel.iNoBays
    sprList.SetText 1, i, Format(gVessel.gBay(i).iHchNo, "00")
    sprList.SetText 2, i, gVessel.gBay(i).sBayNo
    sprList.SetText 3, i, gVessel.gBay(i).sSize & "0"
  Next i
  
End Sub

Private Sub SetVslInfo()
  txtVslCd.Text = gVessel.sVCode
  txtDeckRows.Text = gVessel.iDMaxRows
  txtDeckTiers.Text = gVessel.iDMaxTiers
  txtHoldRows.Text = gVessel.iHMaxRows
  txtHoldTiers.Text = gVessel.iHMaxTiers
  txtHatchCnt.Text = gVessel.iNoHatchs
  txtVslName.Text = gVessel.sVName
  txtCallSign.Text = gVessel.sCallSign
  txtInmarsal.Text = gVessel.sInmarsat
  txtLloyd.Text = gVessel.sLloyd
  txtLoa.Text = IIf(IsNull(gVessel.nLoa), "0", gVessel.nLoa)
  txtLbp.Text = IIf(IsNull(gVessel.nLbp), "0", gVessel.nLbp)
  txtWidth.Text = IIf(IsNull(gVessel.nWidth), "0", gVessel.nWidth)
  txtDepth.Text = IIf(IsNull(gVessel.nDepth), "0", gVessel.nDepth)
  txtTopHgt.Text = IIf(IsNull(gVessel.nTopHgt), "0", gVessel.nTopHgt)
  txtAntHgt.Text = IIf(IsNull(gVessel.nAntHgt), "0", gVessel.nAntHgt)
  txtBgNo.Text = IIf(IsNull(gVessel.iBgNo), "0", gVessel.iBgNo)
  txtBgLength.Text = IIf(IsNull(gVessel.nBgLength), "0", gVessel.nBgLength)
End Sub

Private Sub txtBgLength_GotFocus()
  txtBgLength.SelLength = Len(txtBgLength.Text)
End Sub

Private Sub txtBgNo_GotFocus()
  txtBgNo.SelLength = Len(txtBgNo.Text)
End Sub

Private Sub txtBgNo_LostFocus()
  If Val(txtBgNo.Text) > gVessel.iNoHatchs Or Val(txtBgNo.Text) <= 0 Then
    MsgBox "It can not be larger than Hatch count."
    txtBgNo.Text = gVessel.iBgNo
    txtBgNo.SelLength = Len(txtBgNo.Text)
    txtBgNo.SetFocus
  End If
End Sub

Private Sub txtDeckRows_GotFocus()
  txtDeckRows.SelLength = Len(txtDeckRows.Text)
End Sub

Private Sub txtDeckRows_LostFocus()
  If Val(txtDeckRows.Text) < gVessel.iDMaxRows Then
    MsgBox "It can not be smaller than original value."
    txtDeckRows.Text = gVessel.iDMaxRows
    txtDeckRows.SelLength = Len(txtDeckRows.Text)
    txtDeckRows.SetFocus
  End If
End Sub

Private Sub txtDeckTiers_GotFocus()
  txtDeckTiers.SelLength = Len(txtDeckTiers.Text)
End Sub

Private Sub txtDeckTiers_LostFocus()
  If Val(txtDeckTiers.Text) < gVessel.iDMaxTiers Then
    MsgBox "It can not be smaller than original value."
    txtDeckTiers.Text = gVessel.iDMaxTiers
    txtDeckTiers.SelLength = Len(txtDeckTiers.Text)
    txtDeckTiers.SetFocus
  End If
End Sub

Private Sub txtHatchCnt_GotFocus()
  txtHatchCnt.SelLength = Len(txtHatchCnt.Text)
End Sub

Private Sub txtHatchCnt_LostFocus()
  If Val(txtHatchCnt.Text) < gVessel.iNoHatchs Then
    MsgBox "It can not be smaller than original value."
    txtHatchCnt.Text = gVessel.iNoHatchs
    txtHatchCnt.SelLength = Len(txtHatchCnt.Text)
    txtHatchCnt.SetFocus
  End If
End Sub

Private Sub txtHoldRows_GotFocus()
  txtHoldRows.SelLength = Len(txtHoldRows.Text)
End Sub

Private Sub txtHoldRows_LostFocus()
  If Val(txtHoldRows.Text) < gVessel.iHMaxRows Then
    MsgBox "It can not be smaller than original value."
    txtHoldRows.Text = gVessel.iHMaxRows
    txtHoldRows.SelLength = Len(txtHoldRows.Text)
    txtHoldRows.SetFocus
  End If
End Sub

Private Sub txtHoldTiers_GotFocus()
  txtHoldTiers.SelLength = Len(txtHoldTiers.Text)
End Sub

Private Sub txtHoldTiers_LostFocus()
  If Val(txtHoldTiers.Text) < gVessel.iHMaxTiers Then
    MsgBox "It can not be smaller than original value."
    txtHoldTiers.Text = gVessel.iHMaxTiers
    txtHoldTiers.SelLength = Len(txtHoldTiers.Text)
    txtHoldTiers.SetFocus
  End If
End Sub

Private Sub txtLoa_GotFocus()
  txtLoa.SelLength = Len(txtLoa.Text)
End Sub

Private Sub txtLbp_GotFocus()
  txtLbp.SelLength = Len(txtLbp.Text)
End Sub

Private Sub txtWidth_GotFocus()
  txtWidth.SelLength = Len(txtWidth.Text)
End Sub

Private Sub txtDepth_GotFocus()
  txtDepth.SelLength = Len(txtDepth.Text)
End Sub

Private Sub txtTopHgt_GotFocus()
  txtTopHgt.SelLength = Len(txtTopHgt.Text)
End Sub

Private Sub txtAntHgt_GotFocus()
  txtAntHgt.SelLength = Len(txtAntHgt.Text)
End Sub

Private Sub txtSNo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
      Pic(Index).SetFocus
    Case vbKeyUp
      Call SetSNo(Index)
      Call ReDraw
      
      If gRowTier = "R" Then
        If gHdIdx = 0 And gSnoIdx = gVessel.iHMaxRows Then
          gRowTier = "T"
          gHdIdx = 0: gSnoIdx = 1
        End If
      Else
        If gHdIdx = 1 And gSnoIdx = gVessel.iDMaxTiers Then
          gRowTier = "R"
          gSnoIdx = gVessel.iDMaxRows
        ElseIf gHdIdx = 0 And gSnoIdx = gVessel.iHMaxTiers Then
          gHdIdx = 1: gSnoIdx = 1
        Else
          gSnoIdx = gSnoIdx + 1
        End If
      End If
      
      Call SetSNoTxtBox(Index)
    Case vbKeyDown
      Call SetSNo(Index)
      Call ReDraw
      
      If gRowTier = "R" Then
        If gHdIdx = 1 And gSnoIdx = gVessel.iDMaxRows Then
          gRowTier = "T"
          gHdIdx = 1: gSnoIdx = gVessel.iDMaxTiers
        End If
      Else
        If gHdIdx = 0 And gSnoIdx = 1 Then
          gRowTier = "R"
          gSnoIdx = gVessel.iHMaxRows
        ElseIf gHdIdx = 1 And gSnoIdx = 1 Then
          gHdIdx = 0: gSnoIdx = gVessel.iHMaxTiers
        Else
          gSnoIdx = gSnoIdx - 1
        End If
      End If
      
      Call SetSNoTxtBox(Index)
    Case vbKeyLeft
      Call SetSNo(Index)
      Call ReDraw
      
      If gRowTier = "R" Then
        If gSnoIdx <> 1 Then
          gSnoIdx = gSnoIdx - 1
        End If
      Else
        If gHdIdx = 1 And gSnoIdx = gVessel.iDMaxTiers Then
          gRowTier = "R"
          gSnoIdx = gVessel.iDMaxRows
        ElseIf gHdIdx = 0 And gSnoIdx = 1 Then
          gRowTier = "R"
          gSnoIdx = gVessel.iHMaxRows
        End If
      End If
      
      Call SetSNoTxtBox(Index)
    Case vbKeyRight
      Call SetSNo(Index)
      Call ReDraw
      
      If gRowTier = "R" Then
        If gHdIdx = 1 And gSnoIdx = gVessel.iDMaxRows Then
          gRowTier = "T"
          gSnoIdx = gVessel.iDMaxTiers
        ElseIf gHdIdx = 0 And gSnoIdx = gVessel.iHMaxRows Then
          gRowTier = "T"
          gSnoIdx = 1
        Else
          gSnoIdx = gSnoIdx + 1
        End If
      End If
      
      Call SetSNoTxtBox(Index)
      
    Case vbKeyControl
      Call SetSNo(Index)
      Call ReDraw
      
      If gRowTier = "R" Then
        If gHdIdx = 1 And gSnoIdx = gVessel.iDMaxRows Then
          gRowTier = "T"
          gSnoIdx = gVessel.iDMaxTiers
        Else
          If gHdIdx = 1 Then
            gSnoIdx = gSnoIdx + 1
          Else
            If gSnoIdx <> 1 Then
              gSnoIdx = gSnoIdx - 1
            End If
          End If
        End If
        
      Else
        If gHdIdx = 0 And gSnoIdx = 1 Then
          gRowTier = "R"
          gSnoIdx = gVessel.iHMaxRows
        ElseIf gHdIdx = 1 And gSnoIdx = 1 Then
          gHdIdx = 0: gSnoIdx = gVessel.iHMaxTiers
        Else
          gSnoIdx = gSnoIdx - 1
        End If
      End If
      
      Call SetSNoTxtBox(Index)
      
    Case vbKeyEscape
      Call SetSNo(Index)
      Call ReDraw
      
      If gRowTier = "R" Then
        If gHdIdx = 0 And gSnoIdx = gVessel.iHMaxRows Then
          gRowTier = "T"
          gSnoIdx = 1
        Else
          If gHdIdx = 0 Then
            gSnoIdx = gSnoIdx + 1
          Else
            If gSnoIdx <> 1 Then
              gSnoIdx = gSnoIdx - 1
            End If
          End If
        End If
        
      Else
        If gHdIdx = 1 And gSnoIdx = gVessel.iDMaxTiers Then
          gRowTier = "R"
          gSnoIdx = gVessel.iDMaxRows
        ElseIf gHdIdx = 0 And gSnoIdx = gVessel.iHMaxTiers Then
          gHdIdx = 1: gSnoIdx = 1
        Else
          gSnoIdx = gSnoIdx + 1
        End If
      End If
      
      Call SetSNoTxtBox(Index)
  End Select
End Sub

Private Sub txtNStkWgt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
      Pic(Index).SetFocus
    
    Case vbKeyControl
      Call SetNStkWgt(Index)
      Call ReDraw
      
      If gHdIdx = 1 And gNStkWgtIdx = gVessel.iDMaxRows Then
        gHdIdx = 0
        gNStkWgtIdx = 1
      ElseIf gHdIdx = 0 And gNStkWgtIdx = gVessel.iHMaxRows Then
        gHdIdx = 1
        gNStkWgtIdx = 1
      Else
        gNStkWgtIdx = gNStkWgtIdx + 1
      End If
      
      Call SetNStkWgtTxtBox(Index)
    
    Case vbKeyEscape
      Call SetNStkWgt(Index)
      Call ReDraw
      
      If gHdIdx = 1 And gNStkWgtIdx = 1 Then
        gHdIdx = 0
        gNStkWgtIdx = gVessel.iHMaxRows
      ElseIf gHdIdx = 0 And gNStkWgtIdx = 1 Then
        gHdIdx = 1
        gNStkWgtIdx = gVessel.iDMaxRows
      Else
        gNStkWgtIdx = gNStkWgtIdx - 1
      End If
      
      Call SetNStkWgtTxtBox(Index)
      
  End Select
End Sub

Private Sub SetSNoTxtBox(idx%)
  If gRowTier = "R" Then
    txtSNo(idx).Left = gVessel.gBay(gBayIdx).gRow(gHdIdx, gSnoIdx).iL - 1
    txtSNo(idx).Top = gVessel.gBay(gBayIdx).gRow(gHdIdx, gSnoIdx).iT - 1
    txtSNo(idx).Text = gVessel.gBay(gBayIdx).gRow(gHdIdx, gSnoIdx).sNo
  Else
    txtSNo(idx).Left = gVessel.gBay(gBayIdx).gTier(gHdIdx, gSnoIdx).iL - 1
    txtSNo(idx).Top = gVessel.gBay(gBayIdx).gTier(gHdIdx, gSnoIdx).iT - 1
    txtSNo(idx).Text = gVessel.gBay(gBayIdx).gTier(gHdIdx, gSnoIdx).sNo
  End If
  
  txtSNo(idx).SelLength = Len(txtSNo(idx).Text)
  txtSNo(idx).Visible = True
End Sub

Private Sub SetNStkWgtTxtBox(idx%)
  txtNStkWgt(idx).Left = gVessel.gBay(gBayIdx).gRow(gHdIdx, gNStkWgtIdx).iL_Wgt - 1
  txtNStkWgt(idx).Top = gVessel.gBay(gBayIdx).gRow(gHdIdx, gNStkWgtIdx).iT_Wgt - 1
  txtNStkWgt(idx).Text = Trim(Str(gVessel.gBay(gBayIdx).gRow(gHdIdx, gNStkWgtIdx).nStkWgt))
  
  txtNStkWgt(idx).SelLength = Len(txtNStkWgt(idx).Text)
  txtNStkWgt(idx).Visible = True
End Sub

Private Sub SetSNo(idx%)
  Dim i%, tStr$
  
  tStr = Trim(txtSNo(idx).Text)
  If tStr <> "00" And Val(tStr) = 0 Then tStr = ""
  
  If gRowTier = "R" Then
    gVessel.gBay(gBayIdx).gRow(gHdIdx, gSnoIdx).sNo = IIf(tStr = "", "", Format(Val(tStr), "00"))
    If gSynchro = "H" And gDrawMode <> "B" Then
      For i = 1 To gVessel.iNoBays
        If gHatchIdx = gVessel.gBay(i).iHchNo And gBayIdx <> i Then
          gVessel.gBay(i).gRow(gHdIdx, gSnoIdx).sNo = IIf(tStr = "", "", Format(Val(tStr), "00"))
        End If
      Next i
    End If
  Else
    gVessel.gBay(gBayIdx).gTier(gHdIdx, gSnoIdx).sNo = IIf(tStr = "", "", Format(Val(tStr), "00"))
    If gSynchro = "H" And gDrawMode <> "B" Then
      For i = 1 To gVessel.iNoBays
        If gHatchIdx = gVessel.gBay(i).iHchNo And gBayIdx <> i Then
          gVessel.gBay(i).gTier(gHdIdx, gSnoIdx).sNo = IIf(tStr = "", "", Format(Val(tStr), "00"))
        End If
      Next i
    End If
  End If
End Sub

Private Sub SetNStkWgt(idx%)
  Dim i%, tStr$
  
  tStr = Trim(txtNStkWgt(idx).Text)
  
  gVessel.gBay(gBayIdx).gRow(gHdIdx, gNStkWgtIdx).nStkWgt = IIf(Trim(tStr) = "", 0, Val(tStr))
  If gSynchro = "H" And gDrawMode <> "B" Then
    For i = 1 To gVessel.iNoBays
      If gHatchIdx = gVessel.gBay(i).iHchNo And gBayIdx <> i Then
        gVessel.gBay(i).gRow(gHdIdx, gNStkWgtIdx).nStkWgt = IIf(Trim(tStr) = "", 0, Val(tStr))
      End If
    Next i
  End If
End Sub

Private Sub txtSNo_LostFocus(Index As Integer)
  Call SetSNo(Index)
  
  txtSNo(Index).Visible = False
  Call ReDraw
End Sub

Private Sub txtnStkWgt_LostFocus(Index As Integer)
  Call SetNStkWgt(Index)
  
  txtNStkWgt(Index).Visible = False
  Call ReDraw
End Sub

Private Sub txtVslCd_GotFocus()
  txtVslCd.SelLength = Len(txtVslCd.Text)
End Sub

Private Sub txtVslCd_KeyPress(KeyAscii As Integer)
  If KeyAscii > 96 And KeyAscii < 124 Then             '대문자로 변환
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub txtVslName_GotFocus()
  txtVslName.SelLength = Len(txtVslName.Text)
End Sub

Private Sub txtVslName_KeyPress(KeyAscii As Integer)
  If KeyAscii > 96 And KeyAscii < 124 Then             '대문자로 변환
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub vsVsl_Change()
  picVessel.Top = -vsVsl.Value
  picVessel.SetFocus
End Sub

Private Sub vsVsl_Scroll()
  vsVsl_Change
End Sub

Private Sub hsVsl_Change()
  picVessel.Left = -hsVsl.Value
  picVessel.SetFocus
End Sub

Private Sub hsVsl_Scroll()
  hsVsl_Change
End Sub

Private Sub picBackCenter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    cDX = X
    picCenter.Move picBackCenter.Left, picBackCenter.Top, 4, picBackCenter.Height
    picCenter.Visible = True
  End If
End Sub

Private Sub picBackCenter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    picCenter.Left = picBackCenter.Left + X
  End If
End Sub

Private Sub picBackCenter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim iBasicWidth%
  
  iBasicWidth = 50
  
  If Button = 1 Then
    picCenter.Visible = False
    
    If picBackCenter.Left + (cDX + X) > Me.ScaleWidth - iBasicWidth Then
      picBackCenter.Left = Me.ScaleWidth - picBackCenter.Width - iBasicWidth
      picBack.Left = picBackCenter.Left + picBackCenter.Width
      picBack.Width = iBasicWidth
    ElseIf picBackCenter.Left + (cDX + X) >= iBasicWidth And picBackCenter.Left + (cDX + X) <= Me.ScaleWidth - iBasicWidth Then
      picBackCenter.Left = picBackCenter.Left + (cDX + X)
      picBack.Left = picBackCenter.Left + picBackCenter.Width
      picBack.Width = picBack.Width - (cDX + X)
    Else
      picBackCenter.Left = iBasicWidth
      picBack.Left = picBackCenter.Left + picBackCenter.Width
      picBack.Width = picBack.Width - (cDX + X)
    End If
    
    If tvList.Width + (cDX + X) > Me.ScaleWidth - iBasicWidth Then
      tvList.Width = Me.ScaleWidth - picBackCenter.Width - iBasicWidth
    ElseIf tvList.Width + (cDX + X) >= iBasicWidth And tvList.Width + (cDX + X) <= Me.ScaleWidth - iBasicWidth Then
      tvList.Width = tvList.Width + (cDX + X)
    Else
      tvList.Width = iBasicWidth
    End If
    
    Call SetScroll
  End If
End Sub

Private Sub SetTreeViewItem()
  Dim nodX As Node
  Dim i%, tStr$, sHatch$, sBayNo$
  
  Screen.MousePointer = vbHourglass
  
  tvList.Nodes.Clear
  Set nodX = tvList.Nodes.Add(, , gVessel.sVCode, gVessel.sVCode, "imgVessel")
  Set nodX = tvList.Nodes.Add(gVessel.sVCode, tvwChild, gVessel.sVCode & "S", "Structure", "imgStruct")
  Set nodX = tvList.Nodes.Add(gVessel.sVCode, tvwChild, gVessel.sVCode & "T", "Table", "imgTable")
  Set nodX = tvList.Nodes.Add(gVessel.sVCode & "T", tvwChild, gVessel.sVCode & "T" & "B", "Basic", "imgBasic")
  
  For i = 1 To gVessel.iNoBays
    sHatch = Format(gVessel.gBay(i).iHchNo, "00")
    sBayNo = gVessel.gBay(i).sBayNo
    
    If tStr <> sHatch Then
      Set nodX = tvList.Nodes.Add(gVessel.sVCode & "S", tvwChild, gVessel.sVCode & "S" & sHatch, sHatch & " Hatch", "imgHatch")
      tStr = sHatch
    End If
    
    Set nodX = tvList.Nodes.Add(gVessel.sVCode & "S" & sHatch, tvwChild, gVessel.sVCode & "S" & sHatch & sBayNo, sBayNo & " Bay", "imgBay")
  Next i
  
  For i = tvList.Nodes.Count To 1 Step -1
    tvList.Nodes(i).Expanded = True
  Next i
  
  Screen.MousePointer = vbDefault
End Sub

'셀 색칠하가
'Private Sub SetFillBlockCell(Pic As Control, X%, Y%, bColor As Long, lColor As Long)
'  Pic.DrawWidth = 1
'  Pic.Line (X, Y)-Step(gYcW, gYcH), bColor, BF
'  Pic.Line (X, Y)-Step(gYcW, gYcH), lColor, B
'End Sub

