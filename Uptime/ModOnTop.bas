Attribute VB_Name = "ModOnTop"
'::: Module by X-ICE (Daniel Morgan) :::

':::# Module for "OnTop" form effect so that form it is used on will stay #:::
':::# on top of all other forms and windows #:::

Option Explicit

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub OnTop(Handle As Long)
'Set the form to the top-most layer
SetWindowPos Handle, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
