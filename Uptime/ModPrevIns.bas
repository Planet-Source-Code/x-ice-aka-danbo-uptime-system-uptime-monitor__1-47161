Attribute VB_Name = "ModPrevIns"
'........................................
'Name: IsItRunning
'Type: Public Function
'........................................
Public Function IsItRunning()
'Check if the program is already running, if it is then
'close it because we don't need two of the same program
'running
If App.PrevInstance = True Then
    End
End If
End Function
