Attribute VB_Name = "Module2"
Public red(5000) As Integer
Public green(5000) As Integer
Public blue(5000) As Integer
Public Function readycolor1(First_N As Integer, final_app As Integer, numofcolors As Integer)

Dim s As Integer
Dim val As Integer
Dim test As Integer
Dim U As Integer
Dim v As Integer
Dim w As Integer

val = CInt((1250 * numofcolors) / (final_app - First_N))
If val < 1 Then val = 1
U = 255
w = 255

For s = First_N To final_app

    If v = 0 And U = 255 And w <> 0 Then
        w = w - val
        If w < 0 Then w = 0
    ElseIf U = 255 And w = 0 And v <> 255 Then
        v = v + val
        If v > 255 Then v = 255
    ElseIf v = 255 And w = 0 And U <> 0 Then
        U = U - val
        If U < 0 Then U = 0
    ElseIf U = 0 And v = 255 And w <> 255 Then
        w = w + val
        If w > 255 Then w = 255
    ElseIf w = 255 And U = 0 And v <> 0 Then
        v = v - val
        If v < 0 Then v = 0
    ElseIf w = 255 And v = 0 And U <> 255 Then
        U = U + val
        If U > 255 Then U = 255
    End If

    red(s) = w
    green(s) = v
    blue(s) = U
    
    'If MsgBox(s & vbCr & U & vbCr & v & vbCr & w & vbCr, vbOKCancel) = vbCancel Then Exit Sub
    
Next s
    
End Function

Public Function Click(A1 As Single, B1 As Single, c As Single, d As Single, GAP As Single, maxtries As Integer, axis_length As Single)

Dim p As Single
Dim q As Single
Dim z As Single
Dim zz As Single
    
Dim locx As String
Dim locy As String

Dim cordx As Variant
Dim cordy As Variant
Dim mag As Variant

cordx = A1 + (c * GAP)
locx = Format(cordx, "0.################")
cordy = B1 - (d * GAP)
locy = Format(cordy, "0.################")

Do While z < maxtries And p ^ 2 + q ^ 2 < 4
    zz = p ^ 2 - q ^ 2 + cordx
    q = 2 * p * q + cordy
    p = zz
    z = z + 1
Loop

mag = CLng(2.5 / axis_length)

MsgBox ("Co-ordinates are " & vbCr & "X = " & locx & vbCr & "Y = " & locy & vbCr & "Magnification is " & mag & vbCr & "N = " & z)

End Function

Public Function Shortcolor(n As Integer, J As Integer, K As Integer, numofcolors As Integer) As Integer

If numofcolors = 1 Then
    Shortcolor = 9
    Exit Function
End If

    Dim U As Integer

    U = n Mod numofcolors
    
    Select Case U
        Case 0
            Shortcolor = 9
        Case 1
            Shortcolor = 13
        Case 2
            Shortcolor = 11
        Case Else
            Shortcolor = 1
        End Select

End Function
