VERSION 5.00
Begin VB.Form frmscroll 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "KompSmash - Scroller"
   ClientHeight    =   2985
   ClientLeft      =   3030
   ClientTop       =   1845
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "KompSmashScroller.frx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      Picture         =   "KompSmashScroller.frx":63C8
      ScaleHeight     =   255
      ScaleWidth      =   1575
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "  RE-ENTER ROOM"
         BeginProperty Font 
            Name            =   "Alleycat ICG"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2520
      Picture         =   "KompSmashScroller.frx":7191
      ScaleHeight     =   255
      ScaleWidth      =   1095
      TabIndex        =   14
      Top             =   960
      Width           =   1095
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "   ATTENTION"
         BeginProperty Font 
            Name            =   "Alleycat ICG"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      Picture         =   "KompSmashScroller.frx":7F5A
      ScaleHeight     =   255
      ScaleWidth      =   1095
      TabIndex        =   12
      Top             =   960
      Width           =   1095
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "      NORMAL"
         BeginProperty Font 
            Name            =   "Alleycat ICG"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      Picture         =   "KompSmashScroller.frx":8D23
      ScaleHeight     =   255
      ScaleWidth      =   1095
      TabIndex        =   10
      Top             =   960
      Width           =   1095
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "      IN & OUT"
         BeginProperty Font 
            Name            =   "Alleycat ICG"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3720
      Picture         =   "KompSmashScroller.frx":9AEC
      ScaleHeight     =   255
      ScaleWidth      =   1095
      TabIndex        =   8
      Top             =   960
      Width           =   1095
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "    SPIRAL"
         BeginProperty Font 
            Name            =   "Alleycat ICG"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      Picture         =   "KompSmashScroller.frx":A8B5
      ScaleHeight     =   255
      ScaleWidth      =   2295
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "                     STOP"
         BeginProperty Font 
            Name            =   "Alleycat ICG"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Text            =   "0.55"
      Top             =   2160
      Width           =   645
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "0.9"
      Top             =   2520
      Width           =   645
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFC0C0&
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Text            =   "Text 2 Scroll"
      Top             =   480
      Width           =   4770
   End
   Begin VB.Line Line6 
      X1              =   3000
      X2              =   3000
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Line Line5 
      X1              =   1920
      X2              =   1920
      Y1              =   1200
      Y2              =   1320
   End
   Begin VB.Line Line4 
      X1              =   4200
      X2              =   3600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   4200
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   1320
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   600
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Alleycat ICG"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DESTRUCTION SCROLLER"
      BeginProperty Font 
         Name            =   "Alleycat ICG"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Timeout each line (seconds)"
      BeginProperty Font 
         Name            =   "Alleycat ICG"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   840
      TabIndex        =   5
      Top             =   2160
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(must be in private room)"
      BeginProperty Font 
         Name            =   "Alleycat ICG"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1800
      TabIndex        =   3
      Top             =   1800
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Timeout-4 Line (seconds)"
      BeginProperty Font 
         Name            =   "Alleycat ICG"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   1755
   End
End
Attribute VB_Name = "frmscroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private dStop As Boolean


Private Sub Command2_Click()
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub
If Text3.Text = "" Then Exit Sub
dStop = False
Dim mString As String
Dim GoIn As Boolean
Command1.Enabled = False
COMMAND2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Text1.Locked = True
Command7.Enabled = True
GoIn = True
mString = Text1.Text
Do Until dStop = True
    For i = 1 To 4
        If Len(mString) = Len(Text1.Text) Then
            GoIn = True
        ElseIf Len(mString) = 0 Then
            GoIn = False
        End If
        
        If GoIn = True Then
            mString = Left$(Text1.Text, Len(mString) - 1)
        Else
            mString = Left$(Text1.Text, Len(mString) + 1)
        End If
        ChatSend mString
        If Len(mString) = Len(Text1.Text) Then GoIn = True
        If Len(Text1.Text) = Len(mString) Then GoIn = False
Pause Text2.Text
        If dStop = True Then GoTo Greed
    Next i
    Pause Text3.Text
Loop
Greed:
Command1.Enabled = True
COMMAND2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command6.Enabled = True
Text1.Locked = False
Command7.Enabled = False
ChatSend "-=Ràmþàgè †øølz §¢röllèr=-"
End Sub

Private Sub Command3_Click()
PrivateRoom GetCaption(FindRoom)
End Sub

Private Sub Command4_Click()
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub
If Text3.Text = "" Then Exit Sub
ChatSend "·•· aTTeNTioN  ·•· {S im"
Pause 0.6
ChatSend Text1.Text
Pause 0.6
ChatSend "·•· aTTeNTioN  ·•· {S im"
End Sub

Private Sub Command6_Click()
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub
If Text3.Text = "" Then Exit Sub
dStop = False
Dim mString As String
Command1.Enabled = False
COMMAND2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Text1.Locked = True
Command7.Enabled = True
mString = Text1.Text
Do Until dStop = True
    For i = 1 To 4
        ChatSend mString
Pause Text2.Text
        If dStop = True Then GoTo Greed
    Next i
    Pause Text3.Text
Loop
Greed:
Command1.Enabled = True
COMMAND2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command6.Enabled = True
Text1.Locked = False
Command7.Enabled = False
ChatSend "-=KompSmash §¢röllèr=-"
End Sub

Private Sub Command7_Click()
dStop = True
End Sub

Private Sub Form_Load()
dStop = False


Call FormOnTop(Me.hwnd, True)
End Sub















Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFFFF
Label4.ForeColor = &HFFFFFF
Label5.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
Label7.ForeColor = &HFFFFFF
Picture1.BorderStyle = 0
Picture2.BorderStyle = 0
Picture3.BorderStyle = 0
Picture4.BorderStyle = 0
Picture5.BorderStyle = 0

Label8.ForeColor = &HFFFFFF
End Sub








Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
Picture2.BorderStyle = 0
Picture3.BorderStyle = 0
Picture4.BorderStyle = 0
Picture5.BorderStyle = 1
End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF&
Label4.ForeColor = &HFFFFFF
Label5.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
Label7.ForeColor = &HFFFFFF
End Sub




Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 1
Picture2.BorderStyle = 0
Picture3.BorderStyle = 0
Picture4.BorderStyle = 0
Picture5.BorderStyle = 0
End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF&
Label3.ForeColor = &HFFFFFF
Label5.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
Label7.ForeColor = &HFFFFFF
End Sub




Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
Picture2.BorderStyle = 1
Picture3.BorderStyle = 0
Picture4.BorderStyle = 0
Picture5.BorderStyle = 0
End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFF&
Label4.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
Label7.ForeColor = &HFFFFFF
End Sub




Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
Picture2.BorderStyle = 0
Picture3.BorderStyle = 1
Picture4.BorderStyle = 0
Picture5.BorderStyle = 0
End Sub


Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF&
Label4.ForeColor = &HFFFFFF
Label5.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
Label7.ForeColor = &HFFFFFF
End Sub




Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
Picture2.BorderStyle = 0
Picture3.BorderStyle = 0
Picture4.BorderStyle = 1
Picture5.BorderStyle = 0
End Sub


Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF&
Label4.ForeColor = &HFFFFFF
Label5.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
End Sub





Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFF&
Label7.ForeColor = &HFFFFFF
Label4.ForeColor = &HFFFFFF
Label5.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 1
Picture2.BorderStyle = 0
Picture3.BorderStyle = 0
Picture4.BorderStyle = 0
Picture5.BorderStyle = 0
End Sub




Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
Picture2.BorderStyle = 1
Picture3.BorderStyle = 0
Picture4.BorderStyle = 0
Picture5.BorderStyle = 0
End Sub




Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
Picture2.BorderStyle = 0
Picture3.BorderStyle = 1
Picture4.BorderStyle = 0
Picture5.BorderStyle = 0
End Sub




Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
Picture2.BorderStyle = 0
Picture3.BorderStyle = 0
Picture4.BorderStyle = 1
Picture5.BorderStyle = 0
End Sub




Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
Picture2.BorderStyle = 0
Picture3.BorderStyle = 0
Picture4.BorderStyle = 0
Picture5.BorderStyle = 1
End Sub





Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me

End Sub


Private Sub Label11_Click()
frmscroll.Hide

End Sub


Private Sub Label4_Click()
dStop = True
End Sub

Private Sub Label5_Click()
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub
If Text3.Text = "" Then Exit Sub
dStop = False
Dim mString As String

Text1.Locked = True

mString = Text1.Text
Do Until dStop = True
    For i = 1 To 4
        mString = Right$(mString, 1) & mString
        mString = Left$(mString, Len(mString) - 1)
        ChatSend mString
Pause Text2.Text
        If dStop = True Then GoTo Greed
    Next i
Loop
Greed:

Text1.Locked = False


End Sub

Private Sub Label6_Click()
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub
If Text3.Text = "" Then Exit Sub
dStop = False
Dim mString As String
Dim GoIn As Boolean

Text1.Locked = True

GoIn = True
mString = Text1.Text
Do Until dStop = True
    For i = 1 To 4
        If Len(mString) = Len(Text1.Text) Then
            GoIn = True
        ElseIf Len(mString) = 0 Then
            GoIn = False
        End If
        
        If GoIn = True Then
            mString = Left$(Text1.Text, Len(mString) - 1)
        Else
            mString = Left$(Text1.Text, Len(mString) + 1)
        End If
        ChatSend mString
        If Len(mString) = Len(Text1.Text) Then GoIn = True
        If Len(Text1.Text) = Len(mString) Then GoIn = False
Pause Text2.Text
        If dStop = True Then GoTo Greed
    Next i
    Pause Text3.Text
Loop
Greed:

Text1.Locked = False

End Sub

Private Sub Label7_Click()
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub
If Text3.Text = "" Then Exit Sub
dStop = False
Dim mString As String

Text1.Locked = True

mString = Text1.Text
Do Until dStop = True
    For i = 1 To 4
        ChatSend mString
Pause Text2.Text
        If dStop = True Then GoTo Greed
    Next i
    Pause Text3.Text
Loop
Greed:

Text1.Locked = False

End Sub


Private Sub Label8_Click()
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub
If Text3.Text = "" Then Exit Sub
ChatSend "·•· aTTeNTioN  ·•· {S im"
Pause 0.6
ChatSend Text1.Text
Pause 0.6
ChatSend "·•· aTTeNTioN  ·•· {S im"
End Sub


Private Sub Label9_Click()
PrivateRoom GetCaption(FindRoom)
End Sub


