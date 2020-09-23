<div align="center">

## Toggle a Form's Caption Bar at Runtime


</div>

### Description

This code allows a user to Toggle the caption bar (toolbox and all) on a form at runtime. Normally you cannot visually change some of the form's settings at runtime this solves one of those instances.

Code can be fairly easily changed so as to accept a parameter for if you want the caption or not. This is just a simple example.
 
### More Info
 
This simply changes the window style of the form and then does a couple of extremely quick re-sizes so that the caption area is forced to redraw. This will only work as is if the code is placed in the form you want to affect.

Sub must be called for changes to take effect.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stephen Kent](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephen-kent.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephen-kent-toggle-a-form-s-caption-bar-at-runtime__1-30048/archive/master.zip)

### API Declarations

```
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const SWP_NOMOVE = &H2
Private Const WS_CAPTION = &HC00000
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
```


### Source Code

```
Private Sub ToggleFormCaption()
  Dim rcWindow As RECT
  Dim lRet As Long
  lRet = GetWindowRect(Me.hwnd, rcWindow)
  lRet = SetWindowLong(Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Xor WS_CAPTION)
  lRet = SetWindowPos(Me.hwnd, 0, 0, 0, rcWindow.Right - rcWindow.Left + 1, rcWindow.Bottom - rcWindow.Top, SWP_NOMOVE)
  lRet = SetWindowPos(Me.hwnd, 0, 0, 0, rcWindow.Right - rcWindow.Left, rcWindow.Bottom - rcWindow.Top, SWP_NOMOVE)
End Sub
```

