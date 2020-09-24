VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Uptime"
   ClientHeight    =   630
   ClientLeft      =   12720
   ClientTop       =   495
   ClientWidth     =   2100
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Uptime"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   2100
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrFlash 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer tmrSFlash 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer tmrFUpdate 
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Label lblUp 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "#Load#"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   " System Uptime - How long the computer has been powered up for "
      Top             =   135
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   480
      Left            =   20
      Top             =   15
      Width           =   2025
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount& Lib "kernel32" ()
Dim MouseDownX As Long
Dim MouseDownY As Long

'........................................
'Name: Form_DblClick
'Object: Form
'Event: DblClick(Double click)
'........................................
Private Sub Form_DblClick()
'Unload form(end program) when the form is double clicked on
Unload Me
End Sub

'........................................
'Name: Form_Load
'Object: Form
'Event: Load
'........................................
Private Sub Form_Load()
'Check if the program is already running using a function in the module "ModPrevIns"
IsItRunning
'Set the label caption when the form is loading
lblUp.Caption = "Loading..."
'Size the form according to the size of the box so that the
'line of the box can be seen properly
Me.Height = Shape1.Height + 20
Me.Width = Shape1.Width + 20
OnTop Me.hwnd
'Make the form 70% transparent(it looks better in my opinion)
Call MakeTrans(Me.hwnd, 70 * 255 / 100)
tmrUpdate.Enabled = True
End Sub

'........................................
'Name: lblUp_DblClick
'Object: lblUp
'Event: DblClick(Double click)
'........................................
Private Sub lblUp_DblClick()
'Unload form(end program) when the label is double clicked on
'This allows for unload when you double click anywhere on the form
Unload Me
End Sub

'........................................
'Name: lblUp_MouseDown
'Object: lblUp
'Event: MouseDown
'........................................
Private Sub lblUp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = (Button = vbLeftButton)
'State the mouse co-ordinate variables
MouseDownX = x
MouseDownY = y
End Sub

'........................................
'Name: lblUp_MouseMove
'Object: lblUp
'Event: MouseMove
'........................................
Private Sub lblUp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Set up error handling
On Error GoTo Err:
'Drag the form because it is a borderless form and therefore can't be moved without
'this code
If Button = vbLeftButton Then
    NewX = Me.Left + x - MouseDownX
    If NewX < 0 Then NewX = 0
    If NewX + Me.Width > Screen.Width Then NewX = Screen.Width - Me.Width
    NewY = Me.Top + y - MouseDownY
    If NewY < 0 Then NewY = 0
    If NewY + Me.Height > Screen.Height Then NewY = Screen.Height - Me.Height
    Me.Move NewX + 150, NewY + 150
End If
Exit Sub
Err:
'Display the error and the error description in a message box
MsgBox "Error: " & Err.Description & " occured", vbInformation, " AClock | Error"
End Sub

'........................................
'Name: tmrUpdate_Timer
'Object: tmrUpdate
'Event: Timer
'........................................
Private Sub tmrUpdate_Timer()
'Timer code by VBGUY
Dim Secs, Mins, Hours
Dim TotalMins, TotalHours, TotalSecs, TempSecs
Dim CaptionText As String
Dim StrHours, StrMins, StrSecs As String
   
TotalSecs = Int(GetTickCount / 1000)
TempSecs = Int(Days * 86400)
TotalSecs = TotalSecs - TempSecs
TotalHours = Int((TotalSecs / 60) / 60)
TempSecs = Int(TotalHours * 3600)
TotalSecs = TotalSecs - TempSecs
TotalMins = Int(TotalSecs / 60)
TempSecs = Int(TotalMins * 60)
TotalSecs = (TotalSecs - TempSecs)
If TotalHours > 23 Then
    Hours = (TotalHours - 23)
Else
    Hours = TotalHours
End If
    If TotalMins > 59 Then
        Mins = (TotalMins - (Hours * 60))
    Else
        Mins = TotalMins
    End If
':::# Code by X-ICE#:::

'If Hours(Hours) is less than 10 then add a "0" to the start so it will look like
'"01", "02", "03" and so on
If Hours < 10 Then
    StrHours = "0" & Hours
Else
'Otherwise just use the Hours value
    StrHours = Hours
End If
    
'If Mins(Minutes) is less than 10 then add a "0" to the start so it will look like
'"01", "02", "03" and so on
    If Mins < 10 Then
        StrMins = "0" & Mins
    Else
'Otherwise just use the Mins value
        StrMins = Mins
    End If
'If TotalSecs(Seconds) is less than 10 then add a "0" to the start so it will look like
'"01", "02", "03" and so on
        If TotalSecs < 10 Then
            StrSecs = "0" & TotalSecs
        Else
'Otherwise just use the TotalSecs value
            StrSecs = TotalSecs
        End If
'Create a string from a combination of strings and then use it for the label caption
CaptionText = StrHours & ":" & StrMins & ":" & StrSecs
lblUp.Caption = CaptionText
End Sub

'........................................
'Name: Form_MouseDown
'Object: Form
'Event: MouseDown
'........................................
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'State the mouse co-ordinate variables
MouseDownX = x
MouseDownY = y
End Sub

'........................................
'Name: Form_MouseMove
'Object: Form
'Event: MouseMove
'........................................
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Set up error handling
On Error GoTo Err:
'Drag the form because it is a borderless form and therefore can't be moved without
'this code
If Button = vbLeftButton Then
    NewX = Me.Left + x - MouseDownX
    If NewX < 0 Then NewX = 0
    If NewX + Me.Width > Screen.Width Then NewX = Screen.Width - Me.Width
    NewY = Me.Top + y - MouseDownY
    If NewY < 0 Then NewY = 0
    If NewY + Me.Height > Screen.Height Then NewY = Screen.Height - Me.Height
    Me.Move NewX + 150, NewY + 150
End If
Exit Sub
Err:
'Display the error and the error description in a message box
MsgBox "Error: " & Err.Description & " occured", vbInformation, " AClock | Error"
End Sub

'........................................
'Name: tmrFlash_Timer
'Object: tmrFlash
'Event: Timer
'........................................
Private Sub tmrFlash_Timer()
'Make the form flash by constantly changing the colour of the form and label.
'If the colour is blue then change it to red
If Me.BackColor = vbBlue And lblUp.BackColor = vbBlue Then
    Me.BackColor = vbRed
    lblUp.BackColor = vbRed
'If the colour is blue then change it back to red
Else
    Me.BackColor = vbBlue
    lblUp.BackColor = vbBlue
End If
'# Infinate loops are not a good idea, they are just alot easier in some suitations #
End Sub

'........................................
'Name: tmrSFlash_Timer
'Object: tmrSFlash
'Event: Timer
'........................................
Private Sub tmrSFlash_Timer()
'Stop flashing and reset all properties like backcolor
Me.BackColor = vbBlue
lblUp.BackColor = vbBlue
'Disable timer for flashing form
tmrFlash.Enabled = False
'Disable this timer(the timer to stop flashing)
Me.Enabled = False
End Sub

'........................................
'Name: tmrFUpdate_Timer
'Object: tmrFUpdate
'Event: Timer
'........................................
Private Sub tmrFUpdate_Timer()
'If system uptime is 2 hours then start timers
If lblUp.Caption = "02:00:00" Then
'Timer to make form flash
    tmrFlash.Enabled = True
'Timer to stop form flashing after 5 seconds
    tmrSFlash.Enabled = True
End If
End Sub

