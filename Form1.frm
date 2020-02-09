VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Gallagher Connectivity Tester"
   ClientHeight    =   6675
   ClientLeft      =   6000
   ClientTop       =   2415
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   4755
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1785
      Top             =   7560
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   735
      Top             =   7560
   End
   Begin MSWinsockLib.Winsock Winsock2PaneltoServer 
      Left            =   1260
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.3.194"
      RemotePort      =   1072
      LocalPort       =   1072
   End
   Begin MSWinsockLib.Winsock Winsock1ServertoPanel 
      Left            =   210
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server to Panel"
      Height          =   1485
      Left            =   105
      TabIndex        =   1
      Top             =   5040
      Width           =   4530
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Server Can Talk to Panel"
         Height          =   330
         Left            =   2100
         TabIndex        =   8
         Top             =   420
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status Change"
         Height          =   330
         Left            =   105
         TabIndex        =   6
         Top             =   945
         Width           =   4320
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   645
         Left            =   1995
         Shape           =   4  'Rounded Rectangle
         Top             =   210
         Width           =   2430
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Panel to Server"
      Height          =   2010
      Left            =   105
      TabIndex        =   0
      Top             =   2940
      Width           =   4530
      Begin VB.CommandButton Command2 
         Caption         =   "Test Connection"
         Height          =   540
         Left            =   105
         TabIndex        =   4
         Top             =   735
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1995
         TabIndex        =   3
         Top             =   315
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Panel Can Talk to Server"
         Height          =   330
         Left            =   2205
         TabIndex        =   7
         Top             =   945
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "StatusChange"
         Height          =   330
         Left            =   105
         TabIndex        =   5
         Top             =   1470
         Width           =   4320
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   645
         Left            =   1995
         Shape           =   4  'Rounded Rectangle
         Top             =   735
         Width           =   2430
      End
      Begin VB.Label Label3 
         Caption         =   "Server I.P. Address :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   2
         Top             =   420
         Width           =   1800
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3060
      Left            =   105
      Picture         =   "Form1.frx":0000
      Top             =   105
      Width           =   4560
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      Begin VB.Menu MenuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MenuManual 
      Caption         =   "Manual"
   End
   Begin VB.Menu MEnuHelp 
      Caption         =   "Help"
      Begin VB.Menu MenuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form1.Shape1.BackColor = vbWhite
Form1.Label5.Caption = ""

Winsock1ServertoPanel.Close


Winsock1ServertoPanel.RemoteHost = Text1.Text
Winsock1ServertoPanel.RemotePort = 1072

Form1.Label5.Caption = "Connecting"
Winsock1ServertoPanel.Connect

End Sub

Private Sub Form_Load()

Winsock2PaneltoServer.LocalPort = 1072
Winsock2PaneltoServer.Bind
Winsock2PaneltoServer.Listen

End Sub
Private Sub MenuAbout_Click()
Form2.Show
End Sub

Private Sub MenuExit_Click()
End
End Sub

Private Sub MenuManual_Click()
Form1.Command2.Visible = True
Form1.Text1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Form1.Label5 = Winsock1ServertoPanel.State
If Form1.Label5.Caption = 7 Then Form1.Shape1.BackColor = vbGreen: Form1.Label1.Visible = True: Form1.Label1.Caption = "Panel Can Talk to Server"
If Form1.Label5.Caption = 6 Then Form1.Shape1.BackColor = vbRed: Form1.Label1.Caption = "Panel Can NOT Talk to Server": Form1.Label1.Visible = True

If Form1.Label5.Caption = 8 Then Form1.Shape1.BackColor = vbWhite: Winsock1ServertoPanel.Close

End Sub

Private Sub Timer2_Timer()
Form1.Label6.Caption = Winsock2PaneltoServer.State
If Form1.Label6.Caption = 7 Then Form1.Shape2.BackColor = vbGreen: Form1.Text1.Text = Winsock2PaneltoServer.RemoteHostIP: Form1.Label2.Visible = True

If Form1.Label5.Caption = "7" Then GoTo DONT_CLICK

If Form1.Label6.Caption = 7 Then Command2_Click
DONT_CLICK:

If Form1.Label6.Caption = 8 Then Form1.Shape2.BackColor = vbWhite: Winsock2PaneltoServer.Close: Winsock2PaneltoServer.Listen
End Sub

Private Sub Winsock2PaneltoServer_ConnectionRequest(ByVal requestID As Long)
Winsock2PaneltoServer.Close
Winsock2PaneltoServer.Accept requestID
End Sub

