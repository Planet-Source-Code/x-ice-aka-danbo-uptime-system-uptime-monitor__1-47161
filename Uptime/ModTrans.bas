Attribute VB_Name = "ModTrans"
'::: Module by X-ICE (Daniel Morgan) :::

':::# Module for "Transparent" form effect so that the form it is used on will be
':::# made a transparent window layer. The transparency works by percent
':::# from 1-100 #:::

Option Explicit

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Public Function MakeTrans(ByVal hwnd As Long, TranLev As Integer) As Long
Dim Msg As Long
On Error Resume Next
'Set window style to layered
  Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong hwnd, GWL_EXSTYLE, Msg
'Set the opacity of the layer according to the parameters
  SetLayeredWindowAttributes hwnd, 0, TranLev, LWA_ALPHA
End Function
