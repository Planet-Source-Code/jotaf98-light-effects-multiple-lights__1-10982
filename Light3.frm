VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "''Light III: Multiple Lights'' - Made by Jotaf98"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "Light3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Choose the new background picture..."
      Filter          =   "All supported image files (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg|All files|*.*"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set the background picture..."
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset "
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recomended Colors!"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   6840
      TabIndex        =   16
      Text            =   "10"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6840
      TabIndex        =   14
      Text            =   "20"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6840
      TabIndex        =   11
      Text            =   "50"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2070
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6840
      TabIndex        =   10
      Text            =   "15"
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   120
      ScaleHeight     =   276
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   0
      Top             =   120
      Width           =   6225
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6840
      TabIndex        =   23
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   ";)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1320
      TabIndex        =   19
      Top             =   6450
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      X1              =   3240
      X2              =   3240
      Y1              =   4440
      Y2              =   6600
   End
   Begin VB.Label Label14 
      Caption         =   $"Light3.frx":08CA
      Height          =   975
      Left            =   3480
      TabIndex        =   18
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Label Label13 
      Caption         =   "Blue Brightness:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6480
      TabIndex        =   17
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Green Brightness:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Red Brightness:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6480
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Radius:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   1080
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   3225
      X2              =   3225
      Y1              =   4440
      Y2              =   6600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000016&
      X1              =   3240
      X2              =   120
      Y1              =   5295
      Y2              =   5295
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   3240
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label7 
      Caption         =   "and vote me!"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "http://www.planet-source-code.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "If you liked this effect, please go to"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Then, you can play around with the values! To get a pre-defined color, click ""Recomended Colors!""."
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   $"Light3.frx":09AA
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Label Label8 
      Caption         =   "Made by Jotaf98 (João F. S. Henriques)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label9 
      Caption         =   "E-mail me at: jotaf98@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   5400
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CoolColors(9, 2) As Byte 'Where the "cool colors" will be stored
Dim RndCoolColor As Byte '"Cool color" got at random
Dim BGPictureName As String 'The filename of the background picture

Private Sub RedrawPicture() 'Redraws the picture
    On Error GoTo ErrHandler
    Picture1.Picture = LoadPicture(BGPictureName)
    Exit Sub
    
ErrHandler:
    MsgBox "There was an error displaying the image.", vbExclamation + vbOKOnly
    Picture1.Cls
End Sub

Private Sub Command1_Click()
    RndCoolColor = Rnd * UBound(CoolColors)
    
    MsgBox "The color will now appear in the Red, Green and Blue text fields.", , "Light III: Multiple Lights"
    
    Text4.Text = CoolColors(RndCoolColor, 0)
    Text5.Text = CoolColors(RndCoolColor, 1)
    Text6.Text = CoolColors(RndCoolColor, 2)
End Sub

Private Sub Command2_Click()
    RedrawPicture
End Sub

Private Sub Command3_Click()
    'This will show a dialog box for the user to choose the background image.
    'You'll notice that some properties are already initialized, others are
    'initialized here.
    
    On Error GoTo SkipIt
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.ShowOpen
    BGPictureName = CommonDialog1.FileName
    RedrawPicture
    
SkipIt:
End Sub

Private Sub Form_Load()
    Randomize Timer
    
    Form1.MousePointer = vbHourglass 'Change pointer to hourglass
    Form1.Caption = "''Light III: Multiple Lights'' - Initializing ''Cool Colors'' databank..." 'Change caption
    
    'Initialize the "cool colors" (they're 10)
    InitCoolColor 0, 25, 15, 50
    InitCoolColor 1, 50, 20, 10
    InitCoolColor 2, 15, 40, 5
    InitCoolColor 3, 10, 20, 60
    InitCoolColor 4, 5, 15, 40
    InitCoolColor 5, 60, 15, 35
    InitCoolColor 6, 30, 20, 15
    InitCoolColor 7, 15, 0, 60
    InitCoolColor 8, 60, 15, 0
    InitCoolColor 9, 7, 20, 0
    
    'Redraw the main image
    Form1.Caption = "''Light III: Multiple Lights'' - Redrawing image..." 'Change caption
    BGPictureName = App.Path & "\Geyser.jpg"
    RedrawPicture
    
    Form1.Caption = "''Light III: Multiple Lights'' - Made by Jotaf98" 'Change caption to default
    Form1.MousePointer = vbDefault 'Change pointer to default
End Sub

Private Sub InitCoolColor(Index, RedB, GreenB, BlueB)
    'Initializes a "cool color"
    CoolColors(Index, 0) = RedB
    CoolColors(Index, 1) = GreenB
    CoolColors(Index, 2) = BlueB
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    
    Temp = MsgBox("If you liked this effect, please go to Planet Source Code to vote me ; )" & Chr(13) & Chr(13) & "Visit ''http://www.planet-source-code.com'' now?", vbQuestion + vbYesNo)
    
    If Temp = vbYes Then
        'This code snippet opens an .url file.
        'Please give me some credit if you use it,
        'even if it's just a comment in your
        'program's code ;)
        Shell "rundll32.exe shdocvw.dll,OpenURL " & App.Path & "\Planet Source Code.url"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Sorry, ''Planet Source Code.url'' was not found.", vbExclamation + vbOKOnly
End Sub

Private Sub Label4_Click()
    On Error GoTo ErrHandler
    
    Temp = MsgBox("Visit ''http://www.planet-source-code.com'' now?", vbQuestion + vbYesNo)
    
    If Temp = vbYes Then
        'This code snippet opens an .url file.
        'Please give me some credit if you use it,
        'even if it's just a comment in your
        'program's code ;)
        Shell "rundll32.exe shdocvw.dll,OpenURL " & App.Path & "\Planet Source Code.url"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Sorry, ''Planet Source Code.url'' was not found.", vbExclamation + vbOKOnly
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Make X and Y the coords of the mouse
    Label16.Caption = X
    Label17.Caption = Y
    
    'Draw the light
    DrawLight Picture1.hdc, X, Y, Int(Text4.Text), Int(Text5.Text), _
        Int(Text6.Text), Text3.Text, Int(Text3.Text / 2)
    '                                  ^ I decided to make the number
    '                                    of segments half the radius.
    '                                    Usually higher numbers make
    '                                    better effects, but are slower.
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If the user is clicking, call the MouseDown sub
    If Button = vbLeftButton Then Picture1_MouseDown Button, Shift, X, Y
End Sub

'This is the algorythm for drawing a full circle.
'It's much faster than the Circle function!
'It is provided to you as a bonus; use it as you
'want (you don't need to give me credit for this,
'I found it in one of my Dad's books) ;)
'
'TOP SECRET TIP: If instead of "<=" you use ">="
'(without the quotes -- ""), it will draw only
'outside of the circle. And if instead of
'"If ((X - OrigX..." you use the following
'2 lines, it will draw only the border!
'
'            Temp = ((X - OrigX) * (X - OrigX)) + ((Y - OrigY) * (Y - OrigY))
'            If Temp < (R * R) + R And Temp > (R * R) - R Then
'

Private Sub DrawACircle() 'Just a placeholder, so you don't see it commented
    Dim X As Long 'X Counter
    Dim Y As Long 'Y Counter
    Dim OrigX As Long 'Original X coordinates
    Dim OrigY As Long 'Original Y coordinates
    Dim R As Long 'The circle's radius
    
    OrigX = 100
    OrigY = 200
    R = 25
    
    'Will draw a blue circle in Picture1 at coords:
    'X -> 100 and Y -> 200; with a radius of 25.
    
    For X = OrigX - R To OrigX + R
        For Y = OrigY - R To OrigY + R
            '-- This is the core instruction here --
            If ((X - OrigX) * (X - OrigX)) + ((Y - _
              OrigY) * (Y - OrigY)) <= (R * R) Then
            '----- End of the core instruction -----
                SetPixelV Picture1.hdc, X, Y, vbBlue
            End If
        Next Y
    Next X
End Sub
