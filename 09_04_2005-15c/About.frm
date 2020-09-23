VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3915
   ClientLeft      =   2820
   ClientTop       =   2460
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   3915
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.5c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   5775
   End
   Begin VB.Label lblBackToSystem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back to system"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   492
      Left            =   4800
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By: Kelvin C. Perez Valentin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Text Encryption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   5775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Matrix
    Chr As String
    Color As Long
    End Type
    Dim Matrix(19, 19) As Matrix
    Const RndLetter = "!#$%&0123456789?@ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz~€ƒ†‡‰ŠŒ™šœŸ¡¢£¤¥§©®±µ¶¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ"

Private Sub DrawMatrix(Optional Draw As Boolean = True, Optional Times As Integer = 0)
    Dim A As Integer, B As Integer, C As Integer, D As Integer
    For D = 0 To Times Step (Times = 0) + 1
        For A = 0 To UBound(Matrix, 1)
            For B = 0 To UBound(Matrix, 2)
                If Matrix(A, B).Color = RGB(0, 45, 0) Then Matrix(A, B).Color = RGB(0, 30, 0)
                If Matrix(A, B).Color = 0 Then
                    Matrix(A, B).Chr = ""
                ElseIf Matrix(A, B).Color = RGB(255, 255, 255) Then
                    Matrix(A, B).Color = RGB(0, 255, 0)
                Else
                    Matrix(A, B).Color = Matrix(A, B).Color - RGB(0, 30, 0)
                End If
                If B = 0 Then
                    If Matrix(A, B).Chr = "" And Matrix(A, Int(Rnd * 5)).Chr = "" Then
                        C = Int(Rnd * (Len(RndLetter) + 1) + 1)
                        If C > Len(RndLetter) Then C = Len(RndLetter)
                        Matrix(A, B).Chr = Mid(RndLetter, C, 1)
                        Matrix(A, B).Color = RGB(255, 255, 255)
                    End If
                Else
                    If Matrix(A, B).Chr = "" And Matrix(A, B - 1).Color = RGB(0, 225, 0) Then
                        C = Int(Rnd * (Len(RndLetter) + 1) + 1)
                        If C > Len(RndLetter) Then C = Len(RndLetter)
                            Matrix(A, B).Chr = Mid(RndLetter, C, 1)
                            Matrix(A, B).Color = RGB(0, 255, 0)
                        End If
                End If
                If Draw Then
                    frmAbout.CurrentX = A * 300 + 153 - (frmAbout.TextWidth(Matrix(A, B).Chr) / 2)
                    frmAbout.CurrentY = B * 195 + 15
                    frmAbout.ForeColor = Matrix(A, B).Color
                    frmAbout.Print Matrix(A, B).Chr
                End If
        Next B, A
        If Draw Then DoEvents
    Next
    
End Sub




Private Sub Form_Load()
    'TextH *      #OfLetters       + Boarder
    frmAbout.Height = 195 * (UBound(Matrix, 2) + 1) + 405
    'TextW *       #OfLetters      + Boarder
    frmAbout.Width = 300 * (UBound(Matrix, 1) + 1) + 90
    Randomize
    DrawMatrix False, 100
    Show
    DrawMatrix
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ' -----------------------------------------------------
    ' Quit the from
    ' -----------------------------------------------------
    End
End Sub

Private Sub lblBackToSystem_Click()
    frmAbout.Hide
    'Unload Me
End Sub
