<div align="center">

## Make a Transparent Area \(Any Size\) in your Form


</div>

### Description

This function create a transparent area of dirrent shape (such as rectangle, Circle)

in your form, you specify where and how big the hole is. Unlike most other trnsparant

routine, this one not only let you see trough it, but also allow you total access

access the things in the hole!!! Of course, You can make the entire form transparent

or make you form C - shaped!

Fully tested in VB5 and VB6.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dalin Nie](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dalin-nie.md)
**Level**          |Unknown
**User Rating**    |5.9 (645 globes from 109 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dalin-nie-make-a-transparent-area-any-size-in-your-form__1-1617/archive/master.zip)





### Source Code

```
'1, Declararion
' This should be in the form's General Declaration Area. If you declare in a Modeule,
' you need to omit the word "private"
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
'2 The Function
' This should be in the form's code.
Private Function fMakeATranspArea(AreaType As String, pCordinate() As Long) As Boolean
'Name: fMakeATranpArea
'Author: Dalin Nie
'Date: 5/18/98
'Purpose: Create a Transprarent Area in a form so that you can see through
'Input: Areatype : a String indicate what kind of hole shape it would like to make
' PCordinate : the cordinate area needed for create the shape:
' Example: X1, Y1, X2, Y2 for Rectangle
'OutPut: A boolean
Const RGN_DIFF = 4
Dim lOriginalForm As Long
Dim ltheHole As Long
Dim lNewForm As Long
Dim lFwidth As Single
Dim lFHeight As Single
Dim lborder_width As Single
Dim ltitle_height As Single
 On Error GoTo Trap
 lFwidth = ScaleX(Width, vbTwips, vbPixels)
 lFHeight = ScaleY(Height, vbTwips, vbPixels)
 lOriginalForm = CreateRectRgn(0, 0, lFwidth, lFHeight)
 lborder_width = (lFHeight - ScaleWidth) / 2
 ltitle_height = lFHeight - lborder_width - ScaleHeight
Select Case AreaType
 Case "Elliptic"
 ltheHole = CreateEllipticRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
 Case "RectAngle"
 ltheHole = CreateRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
 Case "RoundRect"
 ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(5), pCordinate(6))
 Case "Circle"
 ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(3), pCordinate(4))
 Case Else
 MsgBox "Unknown Shape!!"
 Exit Function
 End Select
 lNewForm = CreateRectRgn(0, 0, 0, 0)
 CombineRgn lNewForm, lOriginalForm, _
 ltheHole, RGN_DIFF
 SetWindowRgn hWnd, lNewForm, True
 Me.Refresh
 fMakeATranspArea = True
Exit Function
Trap:
 MsgBox "error Occurred. Error # " & Err.Number & ", " & Err.Description
End Function
' 3 How To Call
Dim lParam(1 To 6) As Long
lParam(1) = 100
lParam(2) = 100
lParam(3) = 250
lParam(4) = 250
lParam(5) = 50
lParam(6) = 50
Call fMakeATranspArea("RoundRect", lParam())
'Call fMakeATranspArea("RectAngle", lParam())
'Call fMakeATranspArea("Circle", lParam())
'Call fMakeATranspArea("Elliptic", lParam())
```

