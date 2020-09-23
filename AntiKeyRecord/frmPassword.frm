VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AntiKeyRecord"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1320
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press ENTER to continue."
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -------------------------------------------------------------
' AntiKeyRecord (KeyStroke Recorder Jammer)
' Designed and written by Robin Schuil
' Contact: robin@ykoon.nl
'
' This project demostrates how to prevent keystroke recording
' for your application.
'
' When you decide to implement this method, I would appreciate
' it if you would include my name in the credits.
' -------------------------------------------------------------

' Declare valid password characters
Private Const validChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

' Declare stream data
Private keyBuffer(0 To 99) As Byte
Private keyBufferPtr1 As Integer
Private keyBufferPtr2 As Integer

' Declare password buffer
Private PasswordBuffer As String

Private Sub vbSendKey(KeyAscii As Byte)
    ' Add key to the keyboard buffer
    Text1.SetFocus
    SendKeys Chr(KeyAscii), True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then
        
        ' Backspace was hit
        If Len(PasswordBuffer) > 0 Then PasswordBuffer = Left(PasswordBuffer, Len(PasswordBuffer) - 1)
    
    ElseIf KeyAscii = 13 Then
    
        ' Show password
        MsgBox "The given password is '" & PasswordBuffer & "'.", vbInformation, "KeystrokeJammer"
        End
    
    ElseIf keyBuffer(keyBufferPtr1) <> KeyAscii Then
        
        ' A key was pressed by the user
        If InStr(1, validChars, Chr(KeyAscii)) <> 0 Then
            PasswordBuffer = PasswordBuffer & Chr(KeyAscii)
        End If
        
        ' Reset the randomizer to make it more difficult to predict the sequence
        Randomize Timer
        
    Else
    
        ' key from stream
        keyBufferPtr1 = (keyBufferPtr1 + 1) Mod 100
                
    End If

    ' Cancel event
    KeyAscii = 0
    
    ' Show passwordchars in textfield
    Text1.Text = String(Len(PasswordBuffer), "*")
    Text1.SelStart = Len(PasswordBuffer)

End Sub

Private Sub Form_Load()
    ' Initialize the randomizer
    Randomize Timer
End Sub



Private Sub Timer1_Timer()
        
    Dim myChar As Byte
    Dim i As Integer
    
    Timer1.Enabled = False
    
    ' Select a random character
    i = Int(Rnd * Len(validChars))
    If i = 0 Then i = 1
    myChar = Asc(Mid(validChars, i, 1))
    
    ' Add char to stream and send key
    keyBuffer(keyBufferPtr2) = myChar
    keyBufferPtr2 = (keyBufferPtr2 + 1) Mod 100
    vbSendKey myChar

    DoEvents
    
    ' Get a random interval
    Timer1.Interval = Int(Rnd * 99) + 1
    Timer1.Enabled = True
    
End Sub
