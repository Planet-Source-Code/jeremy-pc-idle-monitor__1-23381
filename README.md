<div align="center">

## PC Idle Monitor


</div>

### Description

This module allows you to determine how long your computer has been Idle. It checks to see if a key was pressed or if the mouse has moved. Well commented.
 
### More Info
 
I hold all copy rights for this code you may distribute it as is. And may not be quoted as your own work so DON'T!

When you call the CheckIdleState Function Make sure you reset the CheckIdleState back to 0 so that It doesn't keep adding to the current value that you used in your code to perform a task.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeremy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeremy.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeremy-pc-idle-monitor__1-23381/archive/master.zip)

### API Declarations

In code below!


### Source Code

```
Option Explicit
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
  X As Long
  Y As Long
End Type
Public Function CheckIdleState()As String
Dim kKey As Integer 'Stores each Key on the keyboard in the for next loop
Dim CurrentMousePos As POINTAPI 'Used to store the current mouse position
Static OldMousePos As POINTAPI 'Static-keeps the old mouse position
Static IdleTime As Date   'Stores the time in a date variable
Dim SystemIdle As Boolean  'Stores weather the systme is idle or not
SystemIdle = True 'Sets the idle value to true
For kKey = 1 To 256 'steps through each key on the keyboard it detect if
 If GetAsyncKeyState(kKey) <> 0 Then 'any of the keys have been pressed
  Debug.Print "Key Pressed"
  SystemIdle = False 'Sets the idle value to false
  Exit For 'Exits the for next loop so that it will move on to the next step
 End If
Next
GetCursorPos CurrentMousePos 'Gets the current cursor position and stores it
If CurrentMousePos.X <> OldMousePos.X Or _
CurrentMousePos.Y <> OldMousePos.Y Then 'Checks to see if the cursor has moved
  Debug.Print "Mouse Moved"
  SystemIdle = False    'since the last time it was checked
End If
OldMousePos = CurrentMousePos 'Stores the current mouse position for comparring positons the
        'next time through
If SystemIdle = True Then 'If a key hasn't been pressed and the mouse hasn't moved
 If DateDiff("s", IdleTime, Now) >= 60 Then 'it sets the return value to the elapsed time value
  IdleTime = Now 'Resets the time to check the next minute for idle
  CheckIdleStaate = CheckIdleState + 1 'sets the return value in minutes of being idle
 End If
Else
 IdleTime = Now 'Sets the new Current Idle Time to check for elapsed time
End If
End Function
```

