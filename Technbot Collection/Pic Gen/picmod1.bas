Attribute VB_Name = "picmod1"
Public Const VK_H = 72
Public Const VK_E = 69
Public Const VK_L = 76
Public Const VK_O = 79
Public Const VK_P = 80
Public Const VK_F = 70
Const KEYEVENTF_KEYUP = &H2
Const INPUT_MOUSE = 0
Const INPUT_KEYBOARD = 1
Const INPUT_HARDWARE = 2
Private Type MOUSEINPUT
  dx As Long
  dy As Long
  mouseData As Long
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type
Private Type KEYBDINPUT
  wVk As Integer
  wScan As Integer
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type
Private Type HARDWAREINPUT
  uMsg As Long
  wParamL As Integer
  wParamH As Integer
End Type
Private Type GENERALINPUT
  dwType As Long
  xi(0 To 23) As Byte
End Type
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Sub Sendkey(bKey As Byte)
    Dim GInput(0 To 1) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    KInput.wVk = bKey  'the key we're going to press
    KInput.dwFlags = 0 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD   ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    'do the same as above, but for releasing the key
    KInput.wVk = bKey  ' the key we're going to realease
    KInput.dwFlags = KEYEVENTF_KEYUP  ' release the key
    GInput(1).dwType = INPUT_KEYBOARD  ' keyboard input
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))
End Sub

