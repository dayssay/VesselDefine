VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Configuration"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5295
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin TabDlg.SSTab sTab 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmConfig.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Database Connection"
         Height          =   2295
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4455
         Begin VB.TextBox txtDbase 
            Height          =   270
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   8
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox txtPass 
            Height          =   270
            IMEMode         =   3  '사용 못함
            Left            =   1800
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   1230
            Width           =   2295
         End
         Begin VB.TextBox txtUId 
            Height          =   270
            Left            =   1800
            TabIndex        =   3
            Top             =   870
            Width           =   2295
         End
         Begin VB.TextBox txtIP 
            Height          =   270
            Left            =   1800
            TabIndex        =   2
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Database :"
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   1710
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Password :"
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   1260
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "UID :"
            Height          =   180
            Left            =   480
            TabIndex        =   6
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Server Address :"
            Height          =   180
            Left            =   360
            TabIndex        =   5
            Top             =   540
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim errFlag As Boolean, f
  
  If Trim(txtIP.Text) = "" Then errFlag = True
  If Trim(txtUId.Text) = "" Then errFlag = True
  If Trim(txtPass.Text) = "" Then errFlag = True
  If Trim(txtDbase.Text) = "" Then errFlag = True
  
  If errFlag Then
    MsgBox "Insert each Value!"
    Exit Sub
  End If
  
  gSvrCfg.SvrIp = txtIP.Text
  gSvrCfg.SvrId = txtUId.Text
  gSvrCfg.SvrPw = txtPass.Text
  gSvrCfg.SvrDb = txtDbase.Text
  
  f = FreeFile
  Open App.Path & "\" & cSvrConfigFile For Output As f
    Write #f, gSvrCfg.SvrIp
    Write #f, gSvrCfg.SvrId
    Write #f, gSvrCfg.SvrPw
    Write #f, gSvrCfg.SvrDb
  Close #f
  
  Call InitBasic
  
  Unload Me
End Sub

Private Sub Form_Load()
  txtIP.Text = gSvrCfg.SvrIp
  txtUId.Text = gSvrCfg.SvrId
  txtPass.Text = gSvrCfg.SvrPw
  txtDbase.Text = gSvrCfg.SvrDb
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mdiVesConfig.Enabled = True
End Sub
