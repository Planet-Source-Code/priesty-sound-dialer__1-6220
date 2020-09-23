VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDialer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialer"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   Icon            =   "frmDialer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Sound"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdDial 
      Caption         =   "Dial"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Phone number..."
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   135
      Width           =   975
   End
   Begin VB.TextBox txtSound 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Sound file..."
      Top             =   120
      Width           =   4575
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select a wav file to open."
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4800
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmDialer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
    On Error GoTo Err
    
    cdOpen.Filter = "Wav Files (*.wav)|*.wav"
    cdOpen.ShowOpen
    txtSound.Text = cdOpen.FileName
    txtNumber.Text = ""
    txtNumber.SetFocus

Err:
    Exit Sub
End Sub

Private Sub cmdDial_Click()
    On Error Resume Next
    
    Dim cBuffer$
    Dim WaitTime As Single, StartTime As Single, FinalTime As Single
    
    WaitTime = 30
    If cmdDial.Caption = "Dial" Then
        cmdDial.Caption = "Hang up"
    
        With MSComm1
            .CommPort = 1
            .Settings = "9600,N,8,1"
            .InputLen = 0
            .PortOpen = True
            .Output = "ATDT" & txtNumber.Text & Chr$(13)
        
            StartTime = Timer
            cmdPlay.Enabled = True
            
            Do
                Me.Caption = Str(Int(Timer - StartTime))
                DoEvents
            Loop Until cmdDial.Caption = "Dial"
            
            Me.Caption = "Dialer"
        End With
    Else
        cmdDial.Caption = "Dial"
        MSComm1.PortOpen = False
        cmdPlay.Enabled = False
    End If
End Sub

Private Sub cSoundPlay(cpath As Variant)
    Variable = sndPlaySound(cpath, SND_ASYNC)
End Sub

Private Sub cSoundStop()
    Variable = sndStopSound(0, SND_ASYNC)
End Sub

Private Sub cmdPlay_Click()
    On Error Resume Next
    
    If cmdPlay.Caption = "Play Sound" Then
        cSoundPlay txtSound.Text
        cmdPlay.Caption = "Stop Sound"
    Else
        cSoundStop
        cmdPlay.Caption = "Play Sound"
    End If
End Sub
