VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim K As Integer
Dim J As Integer
Dim n As Integer
Dim numofcolors As Integer
Dim maxtries As Integer
Dim GAP As Single
Dim A1 As Single
Dim B1 As Single
Dim axis_length As Single

Dim DArray(10) As Single

Private Sub Form_Load()

Dim i As Integer
Dim pixelx As Integer
Dim pixely As Integer
Dim final_app As Integer
Dim First_N As Integer

Dim A As Single
Dim B As Single
Dim X As Single
Dim Y As Single
Dim XX As Single
Dim Data As Single

Dim test As Integer
Dim cred As Integer
Dim cgreen As Integer
Dim cblue As Integer
Dim short As Integer

On Error GoTo DialogError
    
If Form1.Check2 Then
    A = CDbl(A)
    B = CDbl(B)
    X = CDbl(X)
    Y = CDbl(Y)
    XX = CDbl(XX)
    Data = CDbl(Data)

    GAP = CDbl(GAP)
    A1 = CDbl(A1)
    B1 = CDbl(B1)
    axis_length = CDbl(axis_length)
End If

Width = Form1.XRes * 15 + 90
Height = Form1.YRes * 15 + 375
maxtries = Form1.maxtries
numofcolors = Form1.numofcolors
pixelx = Form1.XRes - 1
pixely = Form1.YRes - 1
final_app = Form1.final_app
First_N = Form1.First_N

If Form1.Check3 Then short = 7

If Form1.Option3 Then
    axis_length = 2.5 / Form1.axis_length
Else
    axis_length = Form1.axis_length
End If

GAP = (axis_length / Form1.YRes)

If Form1.Check1 = 1 Then
    A1 = Form1.smallx - ((Form1.XRes / 2) * GAP)
    B1 = Form1.largey + (axis_length / 2)
Else
    A1 = Form1.smallx
    B1 = Form1.largey
End If

A = A1
B = B1

ReDim narray(Form1.XRes, Form1.YRes) As Integer

DArray(0) = A1
DArray(1) = B1
DArray(2) = axis_length
DArray(3) = Form1.maxtries
DArray(4) = Form1.final_app
DArray(5) = Form1.First_N
DArray(6) = Form1.XRes
DArray(7) = Form1.YRes
DArray(8) = numofcolors
DArray(9) = Form1.Check3

i = Module2.readycolor1(First_N, final_app, numofcolors)

Open Form1.filename For Output As #1

For i = 0 To 9
    Data = DArray(i)
    Write #1, Data
Next i

    Show
    
    For K = 0 To (pixelx)
        DoEvents
        For J = 0 To (pixely)
            Do While n < maxtries And X ^ 2 + Y ^ 2 < 4
                XX = X ^ 2 - Y ^ 2 + A: Y = 2 * X * Y + B: X = XX: n = n + 1
            Loop
            
            narray(J, K) = n
            
            Write #1, n
            
            If n = maxtries Then
                PSet (J, K), RGB(0, 0, 0)
            ElseIf n >= final_app Then
                PSet (J, K), RGB(255, 0, 0)
            Else
                
                'test = MsgBox(red(n) & green(n) & blue(n), vbOKCancel)
                'If test = 2 Then Exit Sub
                If Form1.Check3 Then
                    Dim i5 As Integer
                    i5 = Module2.Shortcolor(n, J, K, numofcolors)
                    PSet (J, K), QBColor(i5)
                Else
                    cred = red(n)
                    cgreen = green(n)
                    cblue = blue(n)
                    PSet (J, K), RGB(cred, cgreen, cblue)
                End If
            End If
            
            A = A + GAP
            n = 0
            X = 0
            Y = 0
        Next J
        A = A1
        B = B - GAP
    Next K
Close

Exit Sub

DialogError:
    MsgBox Str(Err.Number) & ":" & Err.Description, , "Error"

End Sub

Private Sub Form_MouseDown(button As Integer, shift As Integer, c As Single, d As Single)
Dim i4 As Integer
i4 = Module2.Click(A1, B1, c, d, GAP, maxtries, axis_length)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close
    Unload Me
End Sub



