VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmLoadVessel 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Load Vessel from Database"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmLoadVessel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7920
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.ListView lstView 
      Height          =   3135
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5530
      View            =   1
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2640
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadVessel.frx":038A
            Key             =   "imgSearch"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadVessel.frx":0924
            Key             =   "imgDel"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadVessel.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadVessel.frx":1458
            Key             =   "imgVslCd"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraBasic 
      BorderStyle     =   0  '없음
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   3600
      Width           =   7815
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   6360
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   1
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
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgSearch"
            Object.ToolTipText     =   "Searching Vessels"
            ImageKey        =   "imgSearch"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgDel"
            Object.ToolTipText     =   "Delete Vessel"
            ImageKey        =   "imgDel"
         EndProperty
      EndProperty
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  '없음
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
         Height          =   225
         Left            =   2040
         TabIndex        =   6
         Text            =   "Select the Vessel which you want to Load or Delete the Vessel"
         Top             =   60
         Width           =   5415
      End
      Begin VB.TextBox txtVslCd 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   30
         MaxLength       =   4
         TabIndex        =   0
         Top             =   30
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmLoadVessel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sVslCd$

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdNext_Click()
  If sVslCd = "" Then
    MsgBox "Please select Vessel Code!"
    Exit Sub
  End If
  
  Call GetVslStructure(sVslCd)
  If gVesFlag Then
    Call SetBasic
  End If
  
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Width = 7920
  Me.Height = 4800
  
  Call SetListViewItem
End Sub

Private Sub SetListViewItem()
  Dim lstItem As ListItems
  Dim i%
  
  Screen.MousePointer = vbHourglass
  
  cmdNext.Enabled = False
  
'  i = 0: lstView.ListItems.Clear
'  SQL = "SELECT SVCODE FROM TB_VD_VESSEL WHERE SVCODE LIKE '" & txtVslCd.Text & "%' ORDER BY SVCODE"
'  Set gRs = G_Host_Con.Execute(SQL)
'  With gRs
'    If Not (.BOF And .EOF) Then
'      Do While Not .EOF
'        i = i + 1
'        lstView.ListItems.Add , , NullTrim(!sVCode), , "imgVslCd"
'        .MoveNext
'      Loop
'    End If
'  End With
  
  i = 0
  lstView.ColumnHeaders.Clear
  lstView.View = lvwReport
  lstView.ColumnHeaders.Add 1, , "Vessel Code"
  lstView.ColumnHeaders.Add 2, , "Vessel Name"
  lstView.ColumnHeaders(2).Width = 3000
  lstView.ColumnHeaders.Add 3, , "CallSign"
  lstView.ColumnHeaders.Add 4, , "최종확인"
  
  SQL = "SELECT SVCODE, SVNAME, SCALLSIGN, INS_ID FROM TB_VD_VESSEL WHERE SVCODE LIKE '" & txtVslCd.Text & "%' ORDER BY SVCODE"
  Set gRs = G_Host_Con.Execute(SQL)
  With gRs
    If Not (.BOF And .EOF) Then
      Do While Not .EOF
        i = i + 1
        
        lstView.ListItems.Add , , NullTrim(!sVCode), , "imgVslCd"
        lstView.ListItems(i).SubItems(1) = NullTrim(!sVName)
        lstView.ListItems(i).SubItems(2) = NullTrim(!sCallSign)
        lstView.ListItems(i).SubItems(3) = NullTrim(!INS_ID)
        .MoveNext
      Loop
    End If
  End With
  gRs.Close
  
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mdiVesConfig.Enabled = True
End Sub

Private Sub lstView_ItemClick(ByVal Item As MSComctlLib.ListItem)
  sVslCd = Item.Text
  cmdNext.Enabled = True
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "imgSearch"
      Call SetListViewItem
    Case "imgDel"
      Call DelVessel
      Call SetListViewItem
  End Select
End Sub

Private Sub txtVslCd_KeyPress(KeyAscii As Integer)
  If KeyAscii > 96 And KeyAscii < 124 Then             '대문자로 변환
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub txtVslCd_KeyUp(KeyCode As Integer, Shift As Integer)
  Call SetListViewItem
End Sub

Private Sub DelVessel()
  If MsgBox("Do you want a delete the vessel definition?", vbQuestion + vbYesNo, "질문") = vbNo Then Exit Sub
  
  Call DelVesselInfo(sVslCd)
End Sub
