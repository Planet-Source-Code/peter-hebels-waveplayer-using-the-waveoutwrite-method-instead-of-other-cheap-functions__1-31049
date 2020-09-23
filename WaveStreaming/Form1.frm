VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "waveOutWrite streaming app"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   5775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   327682
      BorderStyle     =   1
      Enabled         =   0   'False
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   1080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "||"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":04C0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Position:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'Wave-Player project by Peter Hebels, Website "www.phsoft.cjb.net"                        *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

'WARNING do not use the stop button of
'your VisualBasic to stop this program
'Close it by using the X button on the form

'I have made this app using some code found on the net and using C++
'Example files from the MSDN cd
'You can find this examples using the waveOutWrite keyword.

'I have made this app because I hate it to use the other stupid calls
'to play wave files and this code is much more extendible.
'Also this method of playing waves supports larger files then the normal calls
'you can use in VisualBasic.
'Let's just say this is the C++ way to play wave files.

Dim fMovingSlider As Boolean
Dim OnPause As Boolean

Private Sub Command2_Click()
If CD_File.FileName = "" Then
MsgBox "Please open a file by selection open from the file menu", vbInformation
Exit Sub
End If

OnPause = False
Play
Timer1.Enabled = True
Slider1.Enabled = True
End Sub

Private Sub Command3_Click()
On Error GoTo ErrHand
If OnPause = False Then
Timer1.Enabled = False
PausePlay
OnPause = True
Else
OnPause = False
Play
Timer1.Enabled = True
End If

Exit Sub
ErrHand:
End Sub

Private Sub Command4_Click()
If isPlaying = False Then Exit Sub
StopPlay
Slider1.Value = 0
End Sub

Private Sub Form_Load()
Slider1.Min = 0
Slider1.Max = 100
Initialize Me.hwnd
fMovingSlider = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Uncomment the code below if you compile the project to an exe
'Otherwise Visualbasic will crash!!

'StopPlay
'Unload Form1
'End
MsgBox "Thank you for using this app!!" & vbCrLf & "Don't forget to uncomment the Form_Unload code!!" & vbCrLf & "if you compile this project to an exe", vbInformation, "ThanXX"
End Sub

Private Sub MnuExit_Click()
Unload Form1
End Sub

Private Sub MnuOpen_Click()
If isPlaying = True Then
StopPlay
Slider1.Enabled = False
End If

        'I use no OCX files!!
        CD_File.hWndOwner = Me.hwnd
        CD_File.DialogTitle = "Open wave file"
        CD_File.CancelError = False
        
        CD_File.filter = "wave Files (*.wav*)|*.wav*|"
        CD_File.ShowOpen

If Len(CD_File.FileName) = 0 Then
   Exit Sub
End If

Label1.Caption = "Filename: " + CD_File.FileName
OpenFile CD_File.FileName
Slider1.Value = 0
Slider1.Enabled = True
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fMovingSlider = True
Command3_Click
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
fMovingSlider = False
FileSeek (Slider1.Value / 100) * Length
Command2_Click
End Sub

Private Sub Timer1_Timer()
If (fMovingSlider) Then
    Exit Sub
End If
If (Playing() = False) Then
    Timer1.Enabled = False
End If
Slider1.Value = (Position() / Length()) * 100

If Slider1.Value = Slider1.Max Then
Slider1.Value = 0
StopPlay
End If
End Sub
