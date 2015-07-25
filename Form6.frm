VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form6"
   ClientHeight    =   6000
   ClientLeft      =   -555
   ClientTop       =   690
   ClientWidth     =   5985
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sFilename As String

Dim n As Integer
Dim J As Integer
Dim K As Integer
Dim maxtries As Integer
Dim numofcolors As Integer
Dim XRes As Integer
Dim YRes As Integer
Dim final_app As Integer
Dim First_N As Integer

Dim mag As Variant

Dim smallx As Single
Dim largey As Single
Dim axis_length As Single
Dim GAP As Single

Dim DArray(10) As Single
Dim narray() As Integer

Private Sub Form_Load()

On Error GoTo DialogError

Dim pixelx As Integer
Dim pixely As Integer
Dim Data As Integer
Dim i As Integer
Dim test As Integer
Dim cred As Integer
Dim cgreen As Integer
Dim cblue As Integer
Dim short As Integer

Dim message As String

Dim clchange As Variant
Dim fachange As Variant
Dim tychange As Variant
Dim nuchange As Variant

Dim change As Boolean
change = False

Open sFilename For Input As #2
 
For i = 0 To 9
    If EOF(2) Then
        MsgBox "End of File!"
        Close
        Exit Sub
    End If
    Input #2, DArray(i)
    Select Case i
    Case 0
        message = message & "Smallest X" & vbCr & DArray(i) & vbCr
    Case 1
        message = message & "Largest Y" & vbCr & DArray(i) & vbCr
    Case 2
        message = message & "Length of Imaginary Axis" & vbCr & DArray(i) & vbCr
    Case 3
        message = message & "Maximum Iterations" & vbCr & DArray(i) & vbCr
    Case 4
        message = message & "Final Color Apprearance" & vbCr & DArray(i) & vbCr
    Case 5
        message = message & "First color Appearance" & vbCr & DArray(i) & vbCr
    Case 6
        message = message & "Image Width" & vbCr & DArray(i) & vbCr
    Case 7
        message = message & "Image Height" & vbCr & DArray(i) & vbCr
    Case 8
        message = message & "Number of Colors" & vbCr & DArray(i) & vbCr
    Case 9
        message = message & "Short Color?" & vbCr & DArray(i) & vbCr
    End Select
Next i

'MsgBox (message)

smallx = DArray(0)
largey = DArray(1)
axis_length = DArray(2)
maxtries = DArray(3)
final_app = DArray(4)
First_N = DArray(5)
XRes = DArray(6)
YRes = DArray(7)
numofcolors = DArray(8)
short = DArray(9)

pixelx = XRes - 1
pixely = YRes - 1
mag = CLng(2.5 / axis_length)
GAP = axis_length / YRes

Data = MsgBox(message & vbCr & "Magnification is " & mag & vbCr & "Change Color Options?", vbYesNoCancel)

If Data = vbCancel Then
    Unload Me
    Exit Sub
End If

If Data = vbYes Then
    clchange = InputBox("First N?", "Change?", DArray(5))
    If clchange <> "" Then
        First_N = clchange
        change = True
    End If
    
    fachange = InputBox("Final Appearance?", "Change?", DArray(4))
    If fachange <> "" Then
        final_app = fachange
        change = True
    End If
    
    tychange = MsgBox("ShortColor?", vbYesNo, "Change?")
    If tychange = vbYes Then
        short = 1
        change = True
    Else
        short = 0
    End If
    
    nuchange = InputBox("Number of Colors?", "Change?", DArray(8))
    If nuchange <> "" Then
        numofcolors = nuchange
        change = True
    End If

End If

ReDim narray(XRes, YRes)

i = Module2.readycolor1(First_N, final_app, numofcolors)

Width = XRes * 15 + 90
Height = YRes * 15 + 375

Show

    For K = 0 To (pixely)
        DoEvents
        For J = 0 To (pixelx)
        
            If EOF(2) Then
                MsgBox "End of File!"
                Close
                If MsgBox("Save Changes?", vbYesNo) = vbYes Then
                    DArray(4) = final_app
                    DArray(5) = First_N
                    DArray(8) = numofcolors
                    DArray(9) = short
                    savechange
                End If

                Exit Sub
            End If
            
            Input #2, n
            
            narray(J, K) = n
            
            If n = maxtries Then
                PSet (J, K), RGB(0, 0, 0)
            ElseIf n >= final_app Then
                PSet (J, K), RGB(255, 0, 0)
            Else
                
                'test = MsgBox(red(n) & green(n) & blue(n), vbOKCancel)
                'If test = 2 Then Exit Sub
                
                If short = 1 Then
                    Dim i5 As Integer
                    i5 = Module2.Shortcolor(n, J, K, numofcolors)
                    PSet (J, K), QBColor(i5)
                Else
                    cred = Module2.red(n)
                    cgreen = Module2.green(n)
                    cblue = Module2.blue(n)
                    PSet (J, K), RGB(cred, cgreen, cblue)
                End If
            End If
        Next J
    Next K
    Close
If change = True Then
    If MsgBox("Save New Color Scheme?", vbYesNo) = vbNo Then Exit Sub
    DArray(4) = final_app
    DArray(5) = First_N
    DArray(8) = numofcolors
    DArray(9) = short
    savechange
Else
    Exit Sub
End If

Exit Sub
    
DialogError:
        MsgBox Str(Err.Number) & ":" & Err.Description, , "Error"

End Sub

Private Sub Form_MouseDown(button As Integer, shift As Integer, c As Single, d As Single)

Dim i4 As Integer

i4 = Module2.Click(smallx, largey, c, d, GAP, maxtries, axis_length)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close
    Unload Me
End Sub

Public Sub replace(filename As String)

'Close
Dim i2 As Integer
Dim i3 As Integer
Dim data1 As Integer
Dim f As Integer
Dim Data As Integer
Dim message1 As String

Open filename For Output Access Write As #3

DArray(4) = final_app
DArray(5) = First_N

For f = 0 To 9
    Data = DArray(f)
    Write #3, Data
    message1 = message1 & DArray(f) & vbCr
Next f
If MsgBox(message1, vbYesNoCancel) = vbCancel Then Exit Sub
Close

Open filename For Append Access Write As #4
For i2 = 0 To (YRes - 1)
    For i3 = 0 To (XRes - 1)
        data1 = narray(i3, i2)
        Write #4, data1
    Next i3
Next i2
Close
End Sub

Private Sub savechange()

Dim i2 As Integer
Dim i3 As Integer
Dim f As Integer

Dim nfilename As String

    With CommonDialog1
        .filename = sFilename
        .Filter = "All Files (*.*)|*.*"
        .ShowSave
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        nfilename = .filename
    End With

Open nfilename For Output As #3

For f = 0 To 9
    Write #3, DArray(f)
Next f

For i2 = 0 To (YRes - 1)
    For i3 = 0 To (XRes - 1)
        Write #3, narray(i3, i2)
    Next i3
Next i2
End Sub
