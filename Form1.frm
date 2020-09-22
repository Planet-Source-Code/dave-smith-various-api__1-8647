VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1740
   ClientLeft      =   6285
   ClientTop       =   525
   ClientWidth     =   5820
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   5820
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2700
      Top             =   1020
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   600
      Index           =   3
      Left            =   5160
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   540
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Height          =   600
      Index           =   3
      Left            =   4440
      Picture         =   "Form1.frx":10CC
      ScaleHeight     =   540
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   600
      Index           =   2
      Left            =   3720
      Picture         =   "Form1.frx":1E8E
      ScaleHeight     =   540
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Height          =   600
      Index           =   2
      Left            =   3000
      Picture         =   "Form1.frx":2C50
      ScaleHeight     =   540
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Height          =   600
      Index           =   1
      Left            =   1560
      Picture         =   "Form1.frx":3A12
      ScaleHeight     =   540
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   600
      Index           =   1
      Left            =   2280
      Picture         =   "Form1.frx":47D4
      ScaleHeight     =   540
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox back 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   1085
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2050
      Top             =   1020
   End
   Begin VB.PictureBox buffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   600
      Index           =   0
      Left            =   840
      Picture         =   "Form1.frx":5596
      ScaleHeight     =   540
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Height          =   600
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":6358
      ScaleHeight     =   540
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Menu menMenu 
      Caption         =   "Menu"
      Begin VB.Menu memPopUp 
         Caption         =   "Close"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu menAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

AddSysTrayIcon Form1, "Click here for Fly Options."

firsttime = True
buff = Form1.buffer.hdc 'picture buffer

hwnddesk = GetDesktopWindow()
bg = GetWindowDC(hwnddesk)
pback = Form1.back.hdc

posX = 50: posY = 50
picWidth = pic(0).Width         ' globally set the picture widths.
picHeight = pic(0).Height

ind = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim msg As Long
Dim sFilter As String
msg = X / Screen.TwipsPerPixelX
    
Select Case msg
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONUP
    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN
        PopupMenu Form1.menMenu
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
End Select

End Sub

Private Sub memPopUp_Click()

End

End Sub

Private Sub menAbout_Click()

MsgBox "About Box Goes Here!"

End Sub

Private Sub Timer1_Timer()

cycle

End Sub
 
Private Sub Timer2_Timer()

randFly = Int(Rnd * 50)

Select Case randFly                 ' just stop the fly randomly.  Like they do!

Case 25, 4
    Timer1.Enabled = False
Case 5, 10, 15, 20, 30, 35
    Timer1.Enabled = True
End Select

End Sub
