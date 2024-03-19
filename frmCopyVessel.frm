VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmCopyVessel 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Copy Vessel from Database"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmCopyVessel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7920
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   1170
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgSimple"
            Object.ToolTipText     =   "List View"
            ImageKey        =   "imgSimple"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imgDetail"
            Object.ToolTipText     =   "Detail View"
            ImageKey        =   "imgDetail"
         EndProperty
      EndProperty
      Begin VB.TextBox txtVslCd 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   30
         MaxLength       =   4
         TabIndex        =   12
         Top             =   20
         Width           =   855
      End
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
         Left            =   1890
         TabIndex        =   11
         Text            =   "Select the Vessel which you want to Copy"
         Top             =   50
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input the Vessel Code & Name you want to create"
      Height          =   1050
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtVslName 
         Height          =   270
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   1
         Top             =   630
         Width           =   3735
      End
      Begin VB.TextBox txtVessel 
         Height          =   270
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "Vessel Name :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   350
         TabIndex        =   8
         Top             =   675
         Width           =   1470
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "Vessel Code :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   360
         TabIndex        =   7
         Top             =   300
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2280
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopyVessel.frx":038A
            Key             =   "imgSimple"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopyVessel.frx":04E4
            Key             =   "imgDetail"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCopyVessel.frx":063E
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
      Begin VB.CommandButton cmdBack 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5040
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   6360
         TabIndex        =   4
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
   Begin MSComctlLib.ListView lstView 
      Height          =   2055
      Left            =   0
      TabIndex        =   10
      Top             =   1470
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3625
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCopyVessel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sVslCd$
Dim sMode$

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
    gVessel.sVCode = txtVessel.Text
    gVessel.sVName = txtVslName.Text
    
    Call SetBasic
  End If
  
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Width = 7920
  Me.Height = 4800
  
  sMode = "D"
  Call SetListViewItem
End Sub

Private Sub SetListViewItem()
  Dim lstItem As ListItems
  Dim i%
  
  Screen.MousePointer = vbHourglass
  
  cmdNext.Enabled = False
  
  i = 0: lstView.ListItems.Clear
  If sMode = "S" Then
    lstView.View = lvwSmallIcon
    
'    SQL = "SELECT SHIP_CD FROM T_VESSEL WHERE SHIP_CD LIKE '" & txtVslCd.Text & "%' ORDER BY SHIP_CD"
    SQL = "SELECT SVCODE FROM TB_VD_VESSEL WHERE SVCODE LIKE '" & txtVslCd.Text & "%' ORDER BY SVCODE"
    Set gRs = G_Host_Con.Execute(SQL)
    With gRs
      If Not (.BOF And .EOF) Then
        Do While Not .EOF
          lstView.ListItems.Add , , NullTrim(!sVCode), , "imgVslCd"
          .MoveNext
        Loop
      End If
    End With
  Else
    lstView.ColumnHeaders.Clear
    lstView.View = lvwReport
    lstView.ColumnHeaders.Add 1, , "Vessel Code"
    lstView.ColumnHeaders.Add 2, , "Vessel Name"
    lstView.ColumnHeaders(2).Width = 3000
    lstView.ColumnHeaders.Add 3, , "CallSign"
    
'    SQL = "SELECT SHIP_CD, SHIP_NM, CALL_SIGN FROM T_VESSEL WHERE SHIP_CD LIKE '" & txtVslCd.Text & "%' ORDER BY SHIP_CD"
    SQL = "SELECT SVCODE, SVNAME, SCALLSIGN FROM TB_VD_VESSEL WHERE SVCODE LIKE '" & txtVslCd.Text & "%' ORDER BY SVCODE"
    Set gRs = G_Host_Con.Execute(SQL)
    With gRs
      If Not (.BOF And .EOF) Then
        Do While Not .EOF
          i = i + 1
          
          lstView.ListItems.Add , , NullTrim(!sVCode), , "imgVslCd"
          lstView.ListItems(i).SubItems(1) = NullTrim(!sVName)
          lstView.ListItems(i).SubItems(2) = NullTrim(!sCallSign)
          .MoveNext
        Loop
      End If
    End With
  End If
  gRs.Close
  
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mdiVesConfig.Enabled = True
End Sub

Private Sub lstView_ItemClick(ByVal Item As MSComctlLib.ListItem)
  sVslCd = Item.Text
  Call chkFstPhase
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "imgSimple"
      sMode = "S"
      Call SetListViewItem
    Case "imgDetail"
      sMode = "D"
      Call SetListViewItem
  End Select
End Sub

Private Sub txtVessel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call chkFstPhase
End Sub

Private Sub txtVslCd_KeyPress(KeyAscii As Integer)
  If KeyAscii > 96 And KeyAscii < 124 Then             '대문자로 변환
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub txtVslCd_KeyUp(KeyCode As Integer, Shift As Integer)
  Call SetListViewItem
End Sub

Private Sub txtVessel_KeyPress(KeyAscii As Integer)
  If KeyAscii > 96 And KeyAscii < 124 Then             '대문자로 변환
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub txtVessel_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkFstPhase
End Sub

Private Sub txtVslName_KeyPress(KeyAscii As Integer)
  If KeyAscii > 96 And KeyAscii < 124 Then             '대문자로 변환
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub txtVslName_KeyUp(KeyCode As Integer, Shift As Integer)
  Call chkFstPhase
End Sub

Private Sub chkFstPhase()
  If Trim(txtVessel.Text) <> "" And Trim(txtVslName.Text) <> "" Then
    cmdNext.Enabled = True
  Else
    cmdNext.Enabled = False
  End If
End Sub

Private Sub txtVslName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call chkFstPhase
End Sub
