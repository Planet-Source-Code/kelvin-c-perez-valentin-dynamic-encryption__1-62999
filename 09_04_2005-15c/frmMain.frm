VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E39F68&
   Caption         =   "Text Encryption"
   ClientHeight    =   5955
   ClientLeft      =   2820
   ClientTop       =   4395
   ClientWidth     =   9750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin Project1.OsenXPButton OsenXPButton3 
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Quit"
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
      MICON           =   "frmMain.frx":3072
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.OsenXPButton OsenXPButton2 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Decrypt"
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
      MICON           =   "frmMain.frx":308E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.OsenXPButton OsenXPButton1 
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Encrypt"
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
      MICON           =   "frmMain.frx":30AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtEncrypted 
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3480
      Width           =   9255
   End
   Begin VB.TextBox txtDecrypted 
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   9255
   End
   Begin Project1.OsenXPButton OsenXPButton4 
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "About"
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
      MICON           =   "frmMain.frx":30C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypted Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label lblSourcePath 
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypted Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   2985
      Left            =   0
      Picture         =   "frmMain.frx":30E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private WidDiff As Single
Private HeiDiff As Single
Private OldFrmWid As Integer
Private OldFrmHei As Integer

Private Sub Form_Resize()
      Image2.Width = Me.Width
    
Dim CTL As Control
Dim strText As String

      ' the on error is placed because there are some controls
      ' that cant be resized this way..ie combobox, for those items
      ' use the setwindowpos API
      On Error Resume Next

      ' when the form first loads, oldformwid and oldfrmhei will be 0
      If (WidDiff <> 1) Or (HeiDiff <> 1) Then

         ' because these 2 var are declared in the general
         ' section, these var actually store the width and height
         ' of the form the LAST time it was resized.
         ' The difference between the last resize is
         ' the number that we multiply the controls by
         ' to create the controls new width and height
         WidDiff = (Form1.Width / OldFrmWid)
         HeiDiff = (Form1.Height / OldFrmHei)
     
         ' now lets go through all the controls and resize then..eh?
         For Each CTL In Controls
            With CTL
               .Left = (CTL.Left * WidDiff)
               .Top = (CTL.Top * HeiDiff)
               strText = UCase(Trim(.Name))
               If Left(strText, 4) <> "OSEN" Then
                  .Width = (CTL.Width * WidDiff)
                  .Height = (CTL.Height * HeiDiff)
               End If
            End With
         Next CTL
         ' the current forms height and width (after the resize)
         ' are stored for the NEXT resize
         OldFrmWid = Form1.Width
         OldFrmHei = Form1.Height
    
      End If

      Image2.Width = Me.Width
      Me.Refresh
    
End Sub

Private Sub Label1_Click()
      txtEncrypted.Text = ""
      txtEncrypted.Refresh
End Sub

Private Sub lblSourcePath_Click()
      strDecryptedText = ""
      txtDecrypted.Text = strDecryptedText
End Sub

Private Sub OsenXPButton1_Click()
      Call Encrypt_Text
End Sub

Private Sub OsenXPButton2_Click()
    Call Decrypt_Text
End Sub

Private Sub OsenXPButton3_Click()
    End
End Sub

Sub Encrypt_Text()
'
' *************************************************************************
' *
' * Dymanic Encryption
' * Kelvin C. Pérez - Valentín
' * kelvin_perez@msn.com
' *
' * Encrypt text dynamically so the same text will never be encrypted equally
' * without the need to keep track of passwords (but can be modified to supply
' * the password separate from the main string).
' *
' * This is more like a character masking instead of a real ecnryption but will
' * do the trick.
' *
' * You can encrypt the same text as many times as you want and the resulting
' * strings will always be different in size and characters.
' *
' *************************************************************************
' *
' *
Dim strText As String
Dim strEncryptedText As String
Dim strKeyNum As String
Dim strChar1 As String
Dim strChar2 As String
Dim nChar As Integer
Dim nLenght As Integer
Dim nEncKey As Integer
Dim nKeyLenght As Integer
Dim nTextLenght As Integer
Dim nCounter As Integer
    
      strEncryptedText = ""
      
      strText = Trim(txtDecrypted.Text)
      If strText = "" Then
         frmError.lblMsg1 = "Sorry, there's nothing to encrypt."
         frmError.lblMsg2 = "Please, type some text in the encrypt text box and try again."
         frmError.Show (1)
         Exit Sub
      End If
      '
      ' Initialize random-number generator.
      '
      Randomize
      '
      ' Get a random Number between 1 and 100. This will be the multiplier
      ' for the Ascii  value of the characters
      '
      nEncKey = Create_Random_Number(1, 100)
      '
      ' Get a random value betwee 5 and 7. This will be the lenght
      '  (with leading zeros) of the value of the Characters
      '
      nCharSize = Create_Random_Number(5, 10)
      '
      ' Encrypt the Size of the characters and convert it to String.
      ' This size has to be standard so we always get the right character.
      '
      strChar1Size = fEncryptedKeySize(nCharSize)
      '
      ' Convert the KeyNumber to String with leading zeros
      '
      strEncKey = NumToString(nEncKey, nCharSize)
      '
      ' Get the text to encrypt and it's size
      '
      strEncryptedText = ""
      nTextLenght = Len(strText)
      '
      ' Loop thru the text one character at the time
      '
      For nCounter = 1 To nTextLenght
         '
         ' Get the Next Character
         '
         strChar1 = Mid(strText, nCounter, 1)
         '
         ' Get Ascii Value of the character multplied by the Key Number
         '
         nChar = Asc(strChar1) * nEncKey
         '
         ' Get the real size of the string, without the masking characters
         '
         nRealSize = Len(Trim(Str(nChar)))
         '
         ' Get the String version of the Ascii Code with leading zeros
         ' using the Random generated Key Lenght
         '
         strChar2 = NumToString(nChar, nCharSize)
         '
         ' Mix all the chracters together
         '
         strChar1 = Mix_String(strChar2, nRealSize)
         '
         ' Add the Newly generated character to the encrypted text variable
         '
         strEncryptedText = strEncryptedText + strChar1
      Next nCounter
      '
      ' Separate the text in two parts to insert the enc
      ' key in the middle of the string
      '
      nLeft = Len(strEncryptedText) \ 2
      strLeft = Left(strEncryptedText, nLeft)
      
      nRight = Len(strEncryptedText) - nLeft
      strRight = Right(strEncryptedText, nRight)
      '
      ' Add all the strings together to get the final result
      '
      Call InsertInTheMiddle(strEncryptedText, strEncKey)
      Call InsertInTheMiddle(strEncryptedText, strChar1Size)
      '
      ' Add a Dummy string at the begining and end to fool people.
      '
      cDummy = CreateDummy
      strEncryptedText = CreateDummy + strEncryptedText + CreateDummy
      
      txtEncrypted.Text = strEncryptedText
      txtEncrypted.Refresh
    
End Sub


Sub Decrypt_Text()

Dim strText As String
Dim strDecryptedText As String
Dim strKeyNum As String
Dim strChar1 As String
Dim strChar2 As String
Dim nLenght As Integer
Dim nKeyNum As Integer
    
      On Error GoTo ErrorHandler
    
      strTempText = txtEncrypted.Text

      strText = strTempText
      strDecryptedText = ""
      
      
      If strText = "" Then
         frmError.lblMsg1 = "Sorry, there's nothing to Decrypt."
         frmError.lblMsg2 = "Please, type some text in the Decrypt text box and try again."
         frmError.Show (1)
         Exit Sub
      End If
      '
      ' Eliminate the Dummys
      '
      strText = Left(strText, Len(strText) - 4)
      strText = Right(strText, Len(strText) - 4)
      nCharSize = 0
      '
      ' Extract the size of text to decrypt and ecnryption key
      '
      Call Extract_Char_Size(strText, nCharSize)
      Call Extract_Enc_Key(strText, nCharSize, nEncKey)
      '
      ' Decrypt the Size of the encrypted characters
      '
      nTextLenght = Len(strText)
      '
      ' Loop thru text in increments of the Key Size
      '
      For nCounter = 1 To Len(strText) Step nCharSize
         '
         ' Get a Character the size of the key
         '
         strChar1 = Mid(strText, nCounter, nCharSize)
         '
         ' Get the value of the character
         '
         nChar = Remove_Alpha_Chars(strChar1)
         '
         ' Divide the value by the Key to get the real value of the character
         '
         nChar2 = nChar / nEncKey
         '
         ' Convert the value to the character
         '
         strChar2 = Chr(nChar2)
         strDecryptedText = strDecryptedText + strChar2
    
      Next nCounter
      '
      ' Clear any unwanted spaces
      '
      strDecryptedText = Trim(strDecryptedText)
      '
      ' Show the decrypted text
      '
      txtDecrypted.Text = strDecryptedText
      Exit Sub
ErrorHandler:
      If Err.Number <> 0 Then
         frmError.lblMsg1 = "Sorry, there encrypted text doesn't seems to be valid."
         frmError.lblMsg2 = "Please, correct this problem and try again."
         frmError.Show (1)
      Else
         MsgBox (Err.Number)
         Resume
      End If

End Sub

Private Sub OsenXPButton4_Click()
    frmAbout.Show
End Sub


Function NumToString(ByVal nNumber As Integer, ByVal nZeros As Integer) As String
      '
      ' convert a number to string using a fixed size using random letters
      ' in front of the real number to match the desired size
      '
Dim strNumber As String
Dim nLenght As Integer
Dim nCounter As Integer
      '
      ' Check that the zeros to fill are not smaller than the actual size
      '
      strNumber = Trim(Str(nNumber))
      nLenght = Len(strNumber)
      If nZeros < nLenght Then
         nZeros = 0
      End If
      
      nUpperBound = 122
      nLowerBound = 65
      
      For nCounter = 1 To nZeros - nLenght
         '
         ' Ortiginally designed to add a zero in front of the string until
         ' we reach the desired size.
         ' Changed to add random letters (A..z, a..z)
         '
         ' strNumber = "0" + strNumber
         lCreated = False
         Do While lCreated = False
            Randomize
            nNumber = Int((nUpperBound - nLowerBound + 1) * Rnd + nLowerBound)
            If ((nNumber > 90) And (nNumber < 97)) Then
               lCreated = False
            Else
               lCreated = True
            End If
         Loop
         strChar1 = Chr(nNumber)
         strNumber = strChar1 + strNumber
      Next nCounter
      '
      ' return the resulting string
      '
      NumToString = strNumber

End Function

Function CreateDummy() As String
      Randomize
      nUpperBound = 122
      nLowerBound = 48
      For nCounter = 1 To 4
         lCreated = False
         Do While lCreated = False
            nDummy = Int((nUpperBound - nLowerBound + 1) * Rnd + nLowerBound)
            If ((nDummy > 57) And (nDummy < 65)) Or _
            ((nDummy > 90) And (nDummy < 97)) Then
               lCreated = False
            Else
               lCreated = True
            End If
         Loop
         cDummy = cDummy + Chr(nDummy)
      Next nCounter
      CreateDummy = cDummy
End Function


Function fEncryptedKeySize(ByVal nKeySize As Integer) As String

      Randomize
      nLowerBound = 0
      '
      ' Just to fool people....never show the real size in the string
      ' but we need to know what we used in order to decrypt it
      ' so we will store the both in the string but maked.
      '
      nKeyEnc = Int((nKeySize - nLowerBound + 1) * Rnd + nLowerBound)
      nKeySize = nKeySize + nKeyEnc
      '
      ' Return the masked value
      '
      fEncryptedKeySize = NumToString(nKeyEnc, 2) + NumToString(nKeySize, 2)

End Function

Function fDecryptedKeySize(ByVal cKey As String) As Integer
      '
      ' Get the number to decrypt the char size
      '
      nKeySize = Val(Right(strText, 2))
      '
      ' Get the ecnrypted char size
      '
      nKeyEnc = Val(Left(strText, 2))
      '
      ' Get the real char size
      '
      nKeySize = nKeySize - nKeyEnc
      '
      ' Return the real value
      '
      fDecryptedKeySize = nKeySize
End Function


Sub InsertInTheMiddle(ByRef strSourceText, ByVal strTextToInsert)
      '
      ' *************************************************************************
      ' *
      ' * Insert a string in the middle of anither
      ' *
      ' *************************************************************************
      '
      ' Get the half left and half right sides of the text
      '
      nLeft = Len(strSourceText) \ 2
      strLeft = Left(strSourceText, nLeft)
      
      nRight = Len(strSourceText) - nLeft
      strRight = Right(strSourceText, nRight)
      '
      ' Insert strTextToString in the middle of strSourceText
      '
      strSourceText = strLeft + strTextToInsert + strRight

End Sub

Sub Extract_Char_Size(ByRef strText, ByRef nCharSize)
      '
      ' ***********************************************************************
      ' *
      ' * Extract the Character Size from the middle of the exncrypted text
      ' *
      ' ***********************************************************************
      '
      ' Get the half left side of the text
      '
      nLeft = Len(strText) \ 2
      strLeft = Left(strText, nLeft)
      '
      ' Get the half right side of the text
      '
      nRight = Len(strText) - nLeft
      strRight = Right(strText, nRight)
      '
      ' Get the key from the text
      '
      strKeyEnc = Right(strLeft, 2)
      strKeySize = Left(strRight, 2)

      strKeyEnc = Replace_Alpha_Chars(strKeyEnc)
      strKeySize = Replace_Alpha_Chars(strKeySize)

      nKeyEnc = Val(strKeyEnc)
      nKeySize = Val(strKeySize)
      nCharSize = nKeySize - nKeyEnc

      strText = Left(strLeft, Len(strLeft) - 2) + Right(strRight, Len(strRight) - 2)
    
End Sub




Sub Extract_Enc_Key(ByRef strText, ByVal nCharSize, ByRef nEncKey)
      '
      ' ************************************************************************
      ' *
      ' * Extract the Encryption Key from the middle of the encrypted text
      ' *
      ' ************************************************************************
      '
      strEncKey = ""
      '
      ' Get the real size of the text (without the previously
      ' stored character size).
      '
      nLenght = Len(strText) - nCharSize
      '
      ' Get the half left and half right sides of the text
      '
      nLeft = nLenght \ 2
      strLeft = Left(strText, nLeft)
    
      nRight = nLenght - nLeft
      strRight = Right(strText, nRight)
      '
      ' Get the key from the text
      '
      strEncKey = Mid(strText, nLeft + 1, nCharSize)
      strEncKey = Replace_Alpha_Chars(strEncKey)
      '
      ' Get the numeric value of the key
      '
      nEncKey = Val(Trim(strEncKey))
      '
      ' Get the real text to decrypt (left side + right side)
      '
      strText = strLeft + strRight

End Sub




Function Mix_String(ByVal cString As String, ByVal nRealSize As Integer) As String
      '
      ' Mix the alphabetic characters with the numeric characters
      '
      ' Create dynamic arrays to store our data
      '
Dim nCharOrder() As Integer
Dim strChar1Order()
Dim strChar1Pool()
Dim A()
Dim B()
      '
      ' initialize critical variables
      '
      strText = ""
      strTempText = Trim(cString)
      nLowerBound = 1
      nUpperBound = Len(strTempText)
      MaxNumber = nUpperBound ' Must equal the Dim above
      '
      ' Assign a size to our arrays based on the size of our text
      '
      ReDim nCharOrder(nUpperBound)
      ReDim strChar1Order(nUpperBound)
      ReDim strChar1Pool(nUpperBound)
      ReDim A(nUpperBound)
      ReDim B(nUpperBound)
      '
      ' Create a set of Random Sequence Array using Kevin Lawrence's
      ' method found at http://www.planet-source-code.com/vb/scripts/
      ' ShowCode.asp?txtCodeId=892&lngWId=1
      ' ---------------------------------------------------------------------
      ' Set the original array
      '
      For Seq = 1 To nUpperBound
         A(Seq) = Seq
         strChar1Order(Seq) = ""
      Next Seq
      ' Main Loop (mix em all up)
      StartTime = Timer
      Randomize (Timer)
         
      For MainLoop = nUpperBound To nLowerBound Step -1
         ChosenNumber = Int(MainLoop * Rnd + 1)
         B(nUpperBound - MainLoop) = A(ChosenNumber)
         A(ChosenNumber) = A(MainLoop)
      Next MainLoop

      EndTime = Timer
      TotalTime = EndTime - StartTime
      '
      ' End of Kevin's Code
      ' ---------------------------------------------------------------------
      ' Get the alpha Chars of at the begining of the encrypted text and
      ' store them into an array in a random order
      '
      strLeftSide = Left(strTempText, nUpperBound - nRealSize)
    
      For nCounter = 1 To nUpperBound - nRealSize
        
         nCharOrder(nCounter) = B(nCounter)
         strChar1Order(nCharOrder(nCounter)) = Mid(strLeftSide, nCounter, 1)
         ' strText = strText & B(nCounter) & strChar1Order(nCharOrder(nCounter)) & " : "
      
      Next nCounter
      '
      ' Get our encrypted number and set a control variable to extract
      ' the numbers as needed.
      '
      strText = ""
      strRightSide = Right(strTempText, nRealSize)
    
      nControl = 1
      '
      ' fill the remaining spaces with our encrypted number:
      ' Loop thru the array, if the current value is empty,
      ' Append the next character from our encrypted number
      ' to the holding string, else, append the current value.
      '
      For nCounter = 1 To nUpperBound
         If strChar1Order(nCounter) = "" Then
            strChar1Order(nCounter) = Mid(strRightSide, nControl, 1)
            nControl = nControl + 1
         End If
         strText = strText + strChar1Order(nCounter)
      Next nCounter
    
      Mix_String = strText
End Function



Function Create_Random_Number(ByVal nLowerBound, ByVal nUpperBound) As Integer
   '
   ' ******************************************************************
   ' *
   ' * Create a random number between a specified range
   ' *
   ' ******************************************************************
   '
   nRandomNumber = Int((nUpperBound - nLowerBound + 1) * Rnd + nLowerBound)
   Create_Random_Number = nRandomNumber

End Function




Function Remove_Alpha_Chars(strTempText As String) As Integer
   '
   ' ******************************************************************
   ' *
   ' *  Clear the string from unwanted spaces
   ' *
   ' ******************************************************************
   '
      strTempText = Trim(strTempText)
      '
      ' Loop trhu the string, If the current character is numeric,
      ' Append it to the holding string.
      '
      For nCounter = 1 To Len(strTempText)
         strChar1 = Mid(strTempText, nCounter, 1)
         If IsNumeric(strChar1) Then
            strText = strText + strChar1
         End If
      Next nCounter
      '
      ' Get the numeric version of the resulting string and return it's value
      '
      nResult = Val(strText)
      Remove_Alpha_Chars = nResult
End Function

Function Replace_Alpha_Chars(ByVal cString As String) As String

      For nCounter = 1 To Len(cString)
         strChar1 = Mid(cString, nCounter, 1)
         If IsNumeric(strChar1) Then
            strTempString = strTempString + strChar1
         Else
            strTempString = strTempString + "0"
         End If
      Next nCounter
    
      Replace_Alpha_Chars = strTempString
    
End Function




Function Generate_Random_Char(nZeros As Integer) As String

      nUpperBound = 122
      nLowerBound = 65
      
      For nCounter = 1 To nZeros
         '
         ' add a random letter (Lower/uper case) in front of the
         ' string until we reach the desired size
         '
         ' strNumber = "0" + strNumber
         lCreated = False
         Do While lCreated = False
            Randomize
            nNumber = Int((nUpperBound - nLowerBound + 1) * Rnd + nLowerBound)
            If ((nNumber > 90) And (nNumber < 97)) Then
               lCreated = False
            Else
               lCreated = True
            End If
         Loop
         strChar1 = Chr(nNumber)
         Generate_Random_Char = strChar1
      Next nCounter
End Function
