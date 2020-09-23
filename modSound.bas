Attribute VB_Name = "modSound"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0           '  play synchronously (default)
Public Const SND_ASYNC = &H1          '  play asynchronously
Public Const SND_NODEFAULT = &H2      '  silence not default, if sound not found
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_LOOP = &H8           '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_FILENAME = &H20000   '  name is a file name
Public Const SND_RESOURCE = &H40004   '  name is a resource name or atom
