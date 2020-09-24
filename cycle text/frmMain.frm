VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2550
   ClientLeft      =   3585
   ClientTop       =   1950
   ClientWidth     =   1950
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2550
   ScaleWidth      =   1950
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   3375
      Left            =   -120
      TabIndex        =   1
      Top             =   -240
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub pause(duration)
Dim starttime
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop
End Sub

Function intro_shroom_macro()
ascii1 = "                 ¸. ~ -..¸" + Chr(13)
ascii2 = "              ¸·´_;  :'::', `·¸" + Chr(13)
ascii3 = "            .´. †    ::::::::::·." + Chr(13)
ascii4 = "          ,':::)   ø   .:::::::  ';" + Chr(13)
ascii5 = "          ;;·´   __ á    ·······:" + Chr(13)
ascii6 = "         .'    ;´:::::·¸ Ð   ::::: '." + Chr(13)
ascii7 = "       ,´ ¸,,.'--···~ `··-  -..'¸::: `," + Chr(13)
ascii8 = "        `~`~··--·,'    ·.:,'·--··~~´" + Chr(13)
ascii9 = "                 .'     .·::'.   99" + Chr(13)
ascii10 = "              .·:::     ..·:::`:." + Chr(13)
ascii11 = "          `´``´`´`´`´`´`´`´`´`´`´`´`´" + Chr(13)
'Me.Print ascii1 & ascii2 & ascii3 & ascii4 & ascii5 & ascii6 & ascii7 & ascii8 & ascii9 & ascii10 & ascii11
intro_shroom_macro = ascii1 & ascii2 & ascii3 & ascii4 & ascii5 & ascii6 & ascii7 & ascii8 & ascii9 & ascii10 & ascii11

End Function

Private Sub Form_Activate()
Call CycleText("thanks for using u-tech", Label3)
Call CycleText("Cycle Text.", Label4)
Call CycleText(intro_shroom_macro, Label1)
pause 1.75
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = App.Title
End Sub
Private Sub CycleText(Txt As String, lbl1 As Label)
Dim jay
Dim starttime
lbl1.Caption = ""
For jay = 1 To Len(Txt)
    lbl1.Caption = Mid$(Txt, 1, jay)
    starttime = Timer
    Do While Timer - starttime < Val("0.0001")
    DoEvents
    Loop
Next jay
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
