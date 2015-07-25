VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4680
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   3360
      TabIndex        =   30
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   360
      TabIndex        =   28
      Text            =   "1"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   195
      Left            =   3840
      TabIndex        =   27
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3360
      TabIndex        =   23
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   1080
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   2880
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3720
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save As..."
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   360
      TabIndex        =   17
      Text            =   "Mandel"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Text            =   "400"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Text            =   "400"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Text            =   "1"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Text            =   "35"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Text            =   "50"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "2.5"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "1.25"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "-2"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "ShortColor?"
      Height          =   255
      Left            =   3720
      TabIndex        =   31
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "NumOfColors"
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "  Double Precision"
      Height          =   495
      Left            =   3600
      TabIndex        =   26
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Magnification"
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Center"
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "File Name"
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "YRes"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "XRes"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Small N"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Final Appear"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "MaxTries"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Axis Length"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "largeY"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "smallX"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public maxtries As Integer
Public axis_length As Single
Public smallx As Double
Public largey As Double
Public final_app As Integer
Public numofcolors As Integer
Public XRes As Integer
Public YRes As Integer
Public filename As String
Public First_N As Integer

Private Sub Command1_Click()

On Error GoTo DialogError

    smallx = Text1.Text
    largey = Text2.Text
    axis_length = Text3.Text
    maxtries = Text4.Text
    final_app = Text5.Text
    First_N = Text6.Text
    XRes = Text7.Text
    YRes = Text8.Text
    filename = Text9.Text
    numofcolors = Text10.Text
    
    Dim Image As New Form2
    Image.Caption = filename
    Load Image
    Image.Show
    Exit Sub
DialogError:
    MsgBox Str(Err.Number) & ":" & Err.Description, , "Error"
End Sub

Private Sub Command2_Click()
With dlgCommonDialog

        .Filter = "All Files (*.*)|*.*"
        .ShowSave
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        Text9.Text = .filename
    End With
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

