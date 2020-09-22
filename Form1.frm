VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Mission Impossible the Game"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12840
   ForeColor       =   &H00FFFF00&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   856
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   3660
      ScaleHeight     =   1140
      ScaleWidth      =   3030
      TabIndex        =   23
      Top             =   3375
      Visible         =   0   'False
      Width           =   3090
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game Over"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   315
         TabIndex        =   24
         Top             =   300
         Width           =   2475
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      Height          =   2040
      Left            =   2610
      ScaleHeight     =   1980
      ScaleWidth      =   4290
      TabIndex        =   15
      Top             =   5805
      Visible         =   0   'False
      Width           =   4350
      Begin VB.CommandButton Command3 
         Caption         =   "Ok"
         Height          =   405
         Left            =   3330
         TabIndex        =   20
         Top             =   1440
         Width           =   780
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   465
         TabIndex        =   19
         Top             =   1455
         Width           =   2730
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1890
         Left            =   75
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   4185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter your name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   2
         Left            =   465
         TabIndex        =   18
         Top             =   1065
         Width           =   3345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You are the best player "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   1
         Left            =   540
         TabIndex        =   17
         Top             =   660
         Width           =   3345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Congratulation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   0
         Left            =   525
         TabIndex        =   16
         Top             =   90
         Width           =   3375
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   3915
      Left            =   7875
      ScaleHeight     =   3855
      ScaleWidth      =   2850
      TabIndex        =   7
      Top             =   45
      Width           =   2910
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   615
         Index           =   1
         Left            =   1740
         Shape           =   4  'Rounded Rectangle
         Top             =   3150
         Width           =   1050
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1815
         TabIndex        =   29
         Top             =   3435
         Width           =   360
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   1
         Left            =   2535
         TabIndex        =   28
         Top             =   3405
         Width           =   225
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   27
         Top             =   3375
         Width           =   270
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1995
         TabIndex        =   26
         Top             =   3135
         Width           =   555
      End
      Begin VB.Label label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Shadow Off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   105
         TabIndex        =   25
         Top             =   3300
         Width           =   1605
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         Height          =   180
         Left            =   60
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   210
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   3015
         Index           =   0
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   2715
      End
      Begin VB.Label label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Best Player"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   4
         Left            =   720
         TabIndex        =   22
         Top             =   2055
         Width           =   1395
      End
      Begin VB.Label label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   21
         Top             =   2445
         Width           =   1905
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1515
         TabIndex        =   14
         Top             =   1605
         Width           =   1110
      End
      Begin VB.Label label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Best Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   3
         Left            =   1455
         TabIndex        =   13
         Top             =   1215
         Width           =   1245
      End
      Begin VB.Label label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   2
         Left            =   1605
         TabIndex        =   12
         Top             =   525
         Width           =   930
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   225
         TabIndex        =   11
         Top             =   1590
         Width           =   1110
      End
      Begin VB.Label label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   1
         Left            =   420
         TabIndex        =   10
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Energy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   0
         Left            =   330
         TabIndex        =   9
         Top             =   60
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   210
         TabIndex        =   8
         Top             =   480
         Width           =   1110
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   660
         Left            =   1620
         Shape           =   4  'Rounded Rectangle
         Top             =   375
         Width           =   930
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   0
         Left            =   2505
         Shape           =   2  'Oval
         Top             =   3465
         Width           =   255
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   1
         Left            =   2205
         Shape           =   2  'Oval
         Top             =   3465
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   6945
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   190
      TabIndex        =   6
      Top             =   6255
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1470
      Left            =   675
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   5
      Top             =   7050
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   450
      Left            =   2865
      TabIndex        =   4
      Top             =   8115
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   690
      Index           =   1
      Left            =   3225
      TabIndex        =   3
      Top             =   4995
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Height          =   690
      Index           =   0
      Left            =   2190
      TabIndex        =   2
      Top             =   4995
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Draft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7200
      Left            =   7425
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   2025
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   10005
      ScaleHeight     =   183
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   182
      TabIndex        =   0
      Top             =   900
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2460
      Top             =   8040
   End
   Begin VB.Image Image1 
      Height          =   1890
      Left            =   75
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2010
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Play As New clsMP3Player
Private start As Integer

Private Sound_list(0 To 10) As String

Private Shadow As Integer
Private Record_time As String
Private Best_player As String

Private Start_time As Single
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Scale_fator As Single

Private xx As Integer

Private Type point
    x As Integer
    y As Integer
End Type

Private Type Energy_type
    x As Integer
    y As Integer
    Color_point As Long
End Type

Private Energy(0 To 3) As Energy_type

Private Relative_Position(3, 12) As point

Private Roll As String

Private actual_direction As Integer

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Type coor
    x As Integer
    y As Integer
    Nr_Inicial_Resource As Integer
    Nr_Final_Resource As Integer
    AjusteX As Integer
    AjusteY As Integer
    LimiteX As Integer
    LimiteY As Integer
    Frame As Integer
End Type

Private Player_direction(0 To 3) As coor
Private f(0 To 3) As Integer

Private Sub Command1_Click(Index As Integer)

  Dim x As Integer

    Select Case Index
      Case 0
        Roll = Mid$(Roll, 2) + Mid$(Roll, 1, 1)
      Case 1
        Roll = Mid$(Roll, 4) + Mid$(Roll, 1, 3)
    End Select

    x = Val(Mid$(Roll, 1, 1))
    
    Player_direction(x).x = Player_direction(actual_direction).x + (Relative_Position(actual_direction, Player_direction(actual_direction).Frame).x - Relative_Position(x, Player_direction(x).Frame).x) / Scale_fator
    Player_direction(x).y = Player_direction(actual_direction).y + (Relative_Position(actual_direction, Player_direction(actual_direction).Frame).y - Relative_Position(x, Player_direction(x).Frame).y) / Scale_fator

    actual_direction = x

End Sub

Private Sub Command2_Click()

    Picture1 = LoadResPicture(401 + xx, 0)
    xx = xx + 1

End Sub

Private Sub Command3_Click()

  Dim free As Integer

    free = FreeFile
    Open App.Path & "\Record.arq" For Output As free
    Print #1, Label3
    Print #1, Text1
    Close free

    Label4 = Label3
    Label3 = "00:00:00"
    Label1 = 10
    label6 = Text1
    Picture4.Visible = False

    Draw_player

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If Timer1.Enabled = False Then
        Exit Sub
    End If
    
    Select Case KeyCode

      Case 39
        Command1_Click 0
        Draw_player

      Case 37
        Command1_Click 1
        Draw_player

      Case 38
        Draw_player

      Case 40

        Picture2.BackColor = Int(Rnd * &HFFFFFF)
        Draw_player

    End Select

End Sub

Private Sub Form_Load()

  Dim free As Integer
  Dim i As Integer

    Randomize

    Sound_list(0) = App.Path & "\1.wav"
    Sound_list(1) = App.Path & "\2.wav"
    Sound_list(2) = App.Path & "\3.wav"
    Sound_list(3) = App.Path & "\misson impossible.mp3" '"Music"
    Sound_list(4) = App.Path & "\4.wav" 'Game over
    Sound_list(5) = App.Path & "\record.mid"
    Sound_list(6) = App.Path & "\5.wav"
    
    free = FreeFile
    Record_time = "00:00:10"
    If Dir$(App.Path & "\Record.arq") <> "" Then
        Open App.Path & "\Record.arq" For Input As 1
        Input #free, Record_time
        Input #free, Best_player
        Close free
    End If

    'Record_time = "00:00:10"
    Label4 = Record_time
    label6 = Best_player
    Scale_fator = 3
    Label11 = 10 - Scale_fator
    
    Player_direction(0).x = 300
    Player_direction(0).y = 150
    Draft.Width = Draft.Width * 1.5
    Draft.Height = Draft.Height * 1.5

    Relative_Position(0, 1).x = 139: Relative_Position(0, 1).y = 6
    Relative_Position(0, 2).x = 129: Relative_Position(0, 2).y = 11
    Relative_Position(0, 3).x = 120: Relative_Position(0, 3).y = 19
    Relative_Position(0, 4).x = 110: Relative_Position(0, 4).y = 27
    Relative_Position(0, 5).x = 102: Relative_Position(0, 5).y = 30
    Relative_Position(0, 6).x = 91: Relative_Position(0, 6).y = 35
    Relative_Position(0, 7).x = 81: Relative_Position(0, 7).y = 42
    Relative_Position(0, 8).x = 71: Relative_Position(0, 8).y = 49
    Relative_Position(0, 9).x = 60: Relative_Position(0, 9).y = 57
    Relative_Position(0, 10).x = 49: Relative_Position(0, 10).y = 61
    Relative_Position(0, 11).x = 35: Relative_Position(0, 11).y = 70
    Relative_Position(0, 12).x = 22: Relative_Position(0, 12).y = 78

    Relative_Position(1, 1).x = 27: Relative_Position(1, 1).y = 8
    Relative_Position(1, 2).x = 36: Relative_Position(1, 2).y = 13
    Relative_Position(1, 3).x = 46: Relative_Position(1, 3).y = 21
    Relative_Position(1, 4).x = 55: Relative_Position(1, 4).y = 29
    Relative_Position(1, 5).x = 63: Relative_Position(1, 5).y = 34
    Relative_Position(1, 6).x = 74: Relative_Position(1, 6).y = 38
    Relative_Position(1, 7).x = 84: Relative_Position(1, 7).y = 42
    Relative_Position(1, 8).x = 94: Relative_Position(1, 8).y = 48
    Relative_Position(1, 9).x = 105: Relative_Position(1, 9).y = 58
    Relative_Position(1, 10).x = 117: Relative_Position(1, 10).y = 64
    Relative_Position(1, 11).x = 128: Relative_Position(1, 11).y = 72
    Relative_Position(1, 12).x = 141: Relative_Position(1, 12).y = 81

    Relative_Position(2, 1).x = 133: Relative_Position(2, 1).y = 77
    Relative_Position(2, 2).x = 121: Relative_Position(2, 2).y = 69
    Relative_Position(2, 3).x = 111: Relative_Position(2, 3).y = 65
    Relative_Position(2, 4).x = 99: Relative_Position(2, 4).y = 59
    Relative_Position(2, 5).x = 91: Relative_Position(2, 5).y = 53
    Relative_Position(2, 6).x = 81: Relative_Position(2, 6).y = 45
    Relative_Position(2, 7).x = 71: Relative_Position(2, 7).y = 35
    Relative_Position(2, 8).x = 59: Relative_Position(2, 8).y = 33
    Relative_Position(2, 9).x = 51: Relative_Position(2, 9).y = 25
    Relative_Position(2, 10).x = 40: Relative_Position(2, 10).y = 18
    Relative_Position(2, 11).x = 31: Relative_Position(2, 11).y = 13
    Relative_Position(2, 12).x = 21: Relative_Position(2, 12).y = 6

    Relative_Position(3, 1).x = 51: Relative_Position(3, 1).y = 76
    Relative_Position(3, 2).x = 63: Relative_Position(3, 2).y = 69
    Relative_Position(3, 3).x = 74: Relative_Position(3, 3).y = 65
    Relative_Position(3, 4).x = 86: Relative_Position(3, 4).y = 60
    Relative_Position(3, 5).x = 97: Relative_Position(3, 5).y = 51
    Relative_Position(3, 6).x = 106: Relative_Position(3, 6).y = 45
    Relative_Position(3, 7).x = 116: Relative_Position(3, 7).y = 39
    Relative_Position(3, 8).x = 126: Relative_Position(3, 8).y = 32
    Relative_Position(3, 9).x = 136: Relative_Position(3, 9).y = 25
    Relative_Position(3, 10).x = 148: Relative_Position(3, 10).y = 20
    Relative_Position(3, 11).x = 160: Relative_Position(3, 11).y = 14
    Relative_Position(3, 12).x = 174: Relative_Position(3, 12).y = 8

    Player_direction(0).Frame = 1
    Player_direction(1).Frame = 1
    Player_direction(2).Frame = 1
    Player_direction(3).Frame = 1

    Roll = "0231"

    Player_direction(0).Nr_Inicial_Resource = 100
    Player_direction(1).Nr_Inicial_Resource = 200
    Player_direction(2).Nr_Inicial_Resource = 300
    Player_direction(3).Nr_Inicial_Resource = 400

    Player_direction(0).Nr_Final_Resource = 112
    Player_direction(1).Nr_Final_Resource = 212
    Player_direction(2).Nr_Final_Resource = 312
    Player_direction(3).Nr_Final_Resource = 412

    Player_direction(0).AjusteX = -128
    Player_direction(0).AjusteY = 75

    Player_direction(1).AjusteX = 120
    Player_direction(1).AjusteY = 78

    Player_direction(2).AjusteX = -120
    Player_direction(2).AjusteY = -80

    Player_direction(3).AjusteX = 130
    Player_direction(3).AjusteY = -70

    Player_direction(0).LimiteX = Draft.ScaleWidth
    Player_direction(0).LimiteY = Draft.ScaleHeight

    Player_direction(1).LimiteX = Draft.ScaleWidth
    Player_direction(1).LimiteY = Draft.ScaleHeight

    Player_direction(2).LimiteX = Draft.ScaleWidth
    Player_direction(2).LimiteY = Draft.ScaleHeight

    Player_direction(3).LimiteX = Draft.ScaleWidth
    Player_direction(3).LimiteY = Draft.ScaleHeight
    Draw_player
    start = True
    
End Sub

Private Sub Form_Resize()

    Picture3.Move ScaleWidth - Picture3.Width, 0
    Image1.Move 0, 0, ScaleWidth, ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Stop_sounds

End Sub

Private Sub Label2_Click(Index As Integer)

  Dim i As Integer
  Dim x As Integer
  Dim y As Integer

    Stop_sounds

    Picture5.Visible = False
    actual_direction = 0
    Player_direction(0).x = 300
    Player_direction(0).y = 150
    Player_direction(0).Frame = 0
    Label1 = 10
    f(0) = 0
    Timer1.Enabled = True
    Start_time = Time
    For i = 0 To 3
volte:
        x = Int(Rnd * Draft.ScaleWidth)
        y = Int(Rnd * Draft.ScaleWidth)
        If Draft.point(x, y) <> 32768 Then
            GoTo volte
        End If
    
        Energy(i).x = x
        Energy(i).y = y
        Energy(i).Color_point = QBColor(i + 10)
    Next i
    Sound = "Sound" & Trim$(Str(3))
    Play.mp3file = Sound_list(3)
    i = Play.playMP3

End Sub

Private Sub label8_Click()
Shadow = Shadow Xor -1
    If Shadow Then
        label8 = "Shadow On"
        Shape4.BackColor = &HFF
      Else
        label8 = "Shadow Off"
        Shape4.BackColor = &HFFFFFF
    End If
End Sub

Private Sub Label9_Click(Index As Integer)
If Index = 0 Then
    Scale_fator = Scale_fator - 1 * (Scale_fator < 10)
Else
    Scale_fator = Scale_fator + 1 * (Scale_fator > 3)
End If
Label11 = 10 - Scale_fator
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Debug.Print "Relative_Position(3," & xx & ").x="; x & ": Relative_Position(3," & xx & ").y = "; y; ""

End Sub



Private Sub Shadow_Click()

End Sub

Private Sub Timer1_Timer()

    
    Draw_player

End Sub

Private Sub Draw_player()

  Dim i As Integer
  Dim b1 As Integer
  Dim b2 As Integer
  Dim b3 As Integer
  Dim b4 As Integer
  
  Dim k As Integer
  Dim Color_point As Long
  Dim pisca As Integer
   
    
    Draft.Cls

    If Int(Rnd * 2) Then
        pisca = 5
    End If
    
    For i = 0 To 3
        If Energy(i).x Then
            Draft.FillColor = Energy(i).Color_point - pisca
            Draft.Circle (Energy(i).x, Energy(i).y), Int(Rnd * 3 + 2), Energy(i).Color_point - pisca
        End If
    Next i
    Draft.FillColor = 0

    i = actual_direction
    
    b1 = f(i)
    b2 = Player_direction(i).Frame
    b3 = Player_direction(i).x
    b4 = Player_direction(i).y
    
    If f(i) = 0 Then
        f(i) = Player_direction(i).Nr_Inicial_Resource
        Player_direction(i).Frame = 0
    End If
        
    If f(i) = Player_direction(i).Nr_Final_Resource Then
        Player_direction(i).Frame = 0
        f(i) = Player_direction(i).Nr_Inicial_Resource
        Player_direction(i).x = Player_direction(i).x + (Player_direction(i).AjusteX) / Scale_fator
        Player_direction(i).y = Player_direction(i).y + (Player_direction(i).AjusteY) / Scale_fator

        If Player_direction(i).x + Pic.Width < 0 Then
            Player_direction(i).x = Player_direction(i).LimiteX
            Player_direction(i).y = Int(Rnd * Player_direction(i).LimiteY)
        End If
        If Player_direction(i).x > Player_direction(i).LimiteX Then
            Player_direction(i).x = 0
            Player_direction(i).y = Int(Rnd * Player_direction(i).LimiteY)
        End If
        If Player_direction(i).y > Player_direction(i).LimiteY Then
            Player_direction(i).y = 0
            Player_direction(i).x = Int(Rnd * Player_direction(i).LimiteX)
        End If
        If Player_direction(i).y + Pic.Height < 0 Then
            Player_direction(i).y = Player_direction(i).LimiteY
            Player_direction(i).x = Int(Rnd * Player_direction(i).LimiteX)
        End If

    End If

    f(i) = f(i) + 1
    
    If Player_direction(i).Frame = 13 Then
        Player_direction(i).Frame = 1
    End If
        
    Player_direction(i).Frame = Player_direction(i).Frame + 1
        
    Color_point = Draft.point(Player_direction(i).x + Relative_Position(i, Player_direction(i).Frame).x / Scale_fator, Player_direction(i).y + Relative_Position(i, Player_direction(i).Frame).y / Scale_fator + 100 / Scale_fator)
        
    If Color_point <> 32768 Then
            
        For k = 0 To 3
                    
            If Energy(k).Color_point = Color_point Or Energy(k).Color_point = Color_point + pisca Then
                Play_sound (2)
                Energy(k).x = 0
                Label1 = Label1 + 10
                GoTo siga
            End If
        Next k
           
        Play_sound (0)
        Label1 = Label1 - 1
            
        If Label1 = 0 Then
            Stop_sounds
                
            Timer1.Enabled = False
            If Label3 > Label4 Then
                Play_sound (5)
                Picture4.Move ScaleWidth / 2 - Picture4.Width / 2, ScaleHeight / 2 - Picture4.Height / 2
                Picture4.Visible = True
                Text1.SetFocus
                Exit Sub
              Else
                Play_sound (4)
                Picture5.Move ScaleWidth / 2 - Picture5.Width / 2, ScaleHeight / 2 - Picture5.Height / 2
                Picture5.Visible = True
            End If
        End If
            
        f(i) = b1
        Player_direction(i).Frame = b2
        Player_direction(i).x = b3
        Player_direction(i).y = b4
        Command1_Click (Int(Rnd * 1))
        Exit Sub
    End If

siga:
        
    Pic = LoadResPicture(f(i), 0)
    Picture2.Cls

    Call BitBlt(Picture2.hDC, 0, 0, Picture2.Width, Picture2.Height, Pic.hDC, 0, 0, vbMergePaint)
    Call BitBlt(Picture2.hDC, 0, 0, Picture2.Width, Picture2.Height, Pic.hDC, 0, 0, vbSrcAnd)
volte:
         
    TransparentBlt Draft.hDC, Player_direction(i).x, Player_direction(i).y, Pic.ScaleWidth / Scale_fator, Pic.ScaleHeight / Scale_fator, Picture2.hDC, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, 0
        
    If Shadow Then
        Draft.Circle (Player_direction(i).x + Relative_Position(i, Player_direction(i).Frame).x / Scale_fator, Player_direction(i).y + Relative_Position(i, Player_direction(i).Frame).y / Scale_fator + 100 / Scale_fator), 5
    End If
        
    If start Then
        Label3 = Format$(Time - Start_time, "hh:mm:ss")
    Else
    Show_instructions
    End If
    Draft.Refresh
    
    Image1.Picture = Draft.Image
        Sound = "Sound3"
    If Play.IsItPlaying = False And start Then
        Play_sound (3)
    End If
    
End Sub

Private Sub Play_sound(x As Integer)

  Dim r As Long

    Sound = "Sound" & Trim$(Str(x))
    Play.mp3file = Sound_list(x)
    r = Play.playMP3

End Sub

Private Sub Stop_sounds()

  Dim i As Integer

    For i = 0 To 5
        Sound = "Sound" & Trim$(Str(i))
        Play.stopMP3
    Next i

End Sub


Private Sub Show_instructions()

Draft.ForeColor = &HFF
Draft.CurrentX = 100
Draft.CurrentY = 500
Draft.Print "- Use Left/Right Keys to controle the Hero"
Draft.ForeColor = &HFFFF00
Draft.CurrentX = 100
Draft.CurrentY = 540
Draft.Print "- Use Down Key to change the Hero color"
Draft.ForeColor = &HFFFF&
Draft.CurrentX = 100
Draft.CurrentY = 580
Draft.Print "- Each colision to spend energy"
Draft.ForeColor = &HFFFFFF
Draft.CurrentX = 100
Draft.CurrentY = 620
Draft.Print "- To win your time should to be bigger that the best time"
Draft.ForeColor = &HFF0000
Draft.CurrentX = 100
Draft.CurrentY = 660
Draft.Print "- This instructions will to destroy itself after the Start Game  :)"
Draft.ForeColor = 0
End Sub
