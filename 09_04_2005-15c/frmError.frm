VERSION 5.00
Begin VB.Form frmError 
   BackColor       =   &H00E39F68&
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin Project1.OsenXPButton OsenXPButton1 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmError.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMsg2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   10815
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warning!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   -480
      TabIndex        =   1
      Top             =   120
      Width           =   10815
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10815
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   120
      Picture         =   "frmError.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9660
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    Image2.Left = 0
    lblError.Left = 0
    lblMsg1.Left = 0
    lblMsg2.Left = 0

    Image2.Width = Me.Width
    lblError.Width = Me.Width
    lblMsg1.Width = Me.Width
    lblMsg2.Width = Me.Width
End Sub

Private Sub OsenXPButton1_Click()
    Me.Hide
    
End Sub
