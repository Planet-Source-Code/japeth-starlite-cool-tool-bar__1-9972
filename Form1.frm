VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4875
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Tool Bar Design"
      Height          =   4095
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   4695
      Begin VB.CommandButton Command4 
         Caption         =   "Default"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Switch To This"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Switch To This"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Switch To This"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Image DefaultImage 
         Height          =   375
         Left            =   1560
         Picture         =   "Form1.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   120
         Picture         =   "Form1.frx":0DAA
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   120
         Picture         =   "Form1.frx":9C3C
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Image Image3 
         Height          =   405
         Left            =   120
         Picture         =   "Form1.frx":1D6EE
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Form Caption Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Index           =   15
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00800080&
         Height          =   255
         Index           =   14
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Index           =   13
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   12
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H000080FF&
         Height          =   255
         Index           =   11
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H0000FF00&
         Height          =   255
         Index           =   10
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00008080&
         Height          =   255
         Index           =   7
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton ColorBut 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox FormChange 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Image CloseButUp 
      Height          =   210
      Left            =   4600
      Picture         =   "Form1.frx":3D290
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MinButUp 
      Height          =   210
      Left            =   4030
      Picture         =   "Form1.frx":3D572
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MinButDown 
      Height          =   210
      Left            =   4030
      Picture         =   "Form1.frx":3D854
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MinBut 
      Height          =   210
      Left            =   4080
      Picture         =   "Form1.frx":3DB36
      Top             =   30
      Width           =   240
   End
   Begin VB.Image MaxMaxButDown 
      Height          =   210
      Left            =   4320
      Picture         =   "Form1.frx":3DE18
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MaxMaxBut 
      Height          =   210
      Left            =   4320
      Picture         =   "Form1.frx":3E0FA
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MaxButUp 
      Height          =   210
      Left            =   4320
      Picture         =   "Form1.frx":3E3DC
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MaxButDown 
      Height          =   210
      Left            =   4320
      Picture         =   "Form1.frx":3E6BE
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MaxBut 
      Height          =   210
      Left            =   4330
      Picture         =   "Form1.frx":3E9A0
      Top             =   30
      Width           =   240
   End
   Begin VB.Image CloseButDown 
      Height          =   210
      Left            =   4600
      Picture         =   "Form1.frx":3EC82
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image CloseBut 
      Height          =   210
      Left            =   4600
      Picture         =   "Form1.frx":3EF64
      Top             =   30
      Width           =   240
   End
   Begin VB.Label FormCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Caption Just Goes Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   45
      Width           =   3675
   End
   Begin VB.Image iconpic 
      Height          =   245
      Left            =   30
      Picture         =   "Form1.frx":3F246
      Stretch         =   -1  'True
      Top             =   15
      Width           =   245
   End
   Begin VB.Image topbar 
      Height          =   270
      Left            =   0
      Picture         =   "Form1.frx":3F550
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CloseBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CloseBut.Picture = CloseButDown.Picture
End Sub

Private Sub CloseBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CloseBut.Picture = CloseButUp.Picture
Unload Me
End Sub

Private Sub ColorBut_Click(Index As Integer)
FormCaption.ForeColor = ColorBut(Index).BackColor
End Sub

Private Sub Command1_Click()
topbar.Picture = Image1.Picture
End Sub

Private Sub Command2_Click()
topbar.Picture = Image2.Picture
End Sub

Private Sub Command3_Click()
topbar.Picture = Image3.Picture
End Sub

Private Sub Command4_Click()
topbar.Picture = DefaultImage.Picture
End Sub

Private Sub Form_Load()
FormChange.Text = FormCaption.Caption
topbar.Width = Me.Width + 25
CloseBut.Left = Me.Width - 400
MaxBut.Left = Me.Width - 670
MinBut.Left = Me.Width - 920
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
    Me.Caption = ""
Else
    Me.Caption = FormCaption.Caption
End If
topbar.Width = Me.Width + 25
CloseBut.Left = Me.Width - 400
MaxBut.Left = Me.Width - 670
MinBut.Left = Me.Width - 920
End Sub

Private Sub FormCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub FormChange_Change()
FormCaption.Caption = FormChange.Text
End Sub

Private Sub MaxBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.WindowState = vbNormal Then
    MaxBut.Picture = MaxButDown.Picture
Else
    MaxBut.Picture = MaxMaxButDown.Picture
End If
End Sub

Private Sub MaxBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.WindowState = vbNormal Then
    MaxBut.Picture = MaxMaxBut.Picture
    Me.WindowState = vbMaximized
Else
    MaxBut.Picture = MaxButUp.Picture
    Me.WindowState = vbNormal
End If
End Sub

Private Sub MinBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MinBut.Picture = MinButDown.Picture
End Sub

Private Sub MinBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MinBut.Picture = MinButUp.Picture
Me.WindowState = vbMinimized
End Sub

Private Sub topbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
