VERSION 5.00
Begin VB.Form formCenter 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   Picture         =   "formCenter.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrlStep 
      Height          =   255
      Left            =   1980
      Max             =   360
      Min             =   20
      SmallChange     =   20
      TabIndex        =   0
      Top             =   2400
      Value           =   20
      Width           =   1695
   End
   Begin VB.HScrollBar scrlSpeen 
      Height          =   255
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   6
      Top             =   2400
      Value           =   1
      Width           =   1695
   End
   Begin VB.CommandButton cmdShowAll 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1620
      Picture         =   "formCenter.frx":9141
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Open All Windows"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Timer tmrLeft 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   1080
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   3240
   End
   Begin VB.Timer tmrRight 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   1200
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   480
   End
   Begin VB.CommandButton cmdShowRight 
      Height          =   615
      Left            =   3000
      Picture         =   "formCenter.frx":9A0B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdShowLeft 
      Height          =   615
      Left            =   240
      Picture         =   "formCenter.frx":9DF5
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdShowDown 
      Height          =   615
      Left            =   1620
      Picture         =   "formCenter.frx":A1EC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdShowUp 
      Height          =   615
      Left            =   1620
      Picture         =   "formCenter.frx":A5F2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblStep 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   2760
      Width           =   615
   End
   Begin VB.Shape shpShadow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   6
      Left            =   2040
      Top             =   2460
      Width           =   1695
   End
   Begin VB.Label lblStpesT 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Steps:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1980
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblSpeedT 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.Shape shpShadow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   5
      Left            =   180
      Top             =   2460
      Width           =   1695
   End
   Begin VB.Shape shpShadow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Index           =   4
      Left            =   1680
      Top             =   1620
      Width           =   615
   End
   Begin VB.Shape shpShadow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Index           =   3
      Left            =   1680
      Top             =   3060
      Width           =   615
   End
   Begin VB.Shape shpShadow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Index           =   2
      Left            =   3060
      Top             =   1620
      Width           =   615
   End
   Begin VB.Shape shpShadow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Index           =   1
      Left            =   300
      Top             =   1620
      Width           =   615
   End
   Begin VB.Shape shpShadow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Index           =   0
      Left            =   1680
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "formCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private formOpened(4) As Boolean

' Open / close all windows
Private Sub cmdShowAll_Click()
    cmdShowDown_Click
    cmdShowLeft_Click
    cmdShowRight_Click
    cmdShowUp_Click
End Sub

' Open lower menu button
Private Sub cmdShowDown_Click()
    If formOpened(2) = False Then
        cmdShowDown.Enabled = False
        formDown.Top = formCenter.Top
        formDown.Left = formCenter.Left
        formDown.Show
        formCenter.Show
    End If
    tmrDown = True
End Sub

Private Sub cmdShowLeft_Click()
    If formOpened(3) = False Then
        cmdShowLeft.Enabled = False
        formLeft.Top = formCenter.Top
        formLeft.Left = formCenter.Left
        formLeft.Show
        formCenter.Show
    End If
    tmrLeft = True
End Sub

' Open right menu
Private Sub cmdShowRight_Click()
    If formOpened(1) = False Then
        cmdShowRight.Enabled = False
        formRight.Top = formCenter.Top
        formRight.Left = formCenter.Left
        formRight.Show
        formCenter.Show
    End If
    tmrRight = True
End Sub

' Open Upper menu
Private Sub cmdShowUp_Click()
    If formOpened(0) = False Then
        cmdShowUp.Enabled = False
        formUp.Top = formCenter.Top
        formUp.Left = formCenter.Left
        formUp.Show
        formCenter.Show
    End If
    tmrUp = True
End Sub

' form load
Private Sub Form_Load()
    Dim i As Long
    For i = 0 To 3
        formOpened(i) = False
    Next
End Sub

Private Sub scrlSpeen_Change()
    tmrDown.Interval = scrlSpeen.Value
    tmrLeft.Interval = scrlSpeen.Value
    tmrRight.Interval = scrlSpeen.Value
    tmrUp.Interval = scrlSpeen.Value
    lblSpeed = scrlSpeen.Value
End Sub

' step size scroller
Private Sub scrlStep_Change()
    lblStep = scrlStep.Value
End Sub

' Timer - Down
Private Sub tmrDown_Timer()
    With formDown
        If formOpened(2) = False Then
            .Top = .Top + lblStep
            If .Top >= formCenter.Top + formCenter.Height Then
                tmrDown = False
                formOpened(2) = True
                cmdShowDown.Enabled = True
            End If
        Else
            .Top = .Top - lblStep
            If .Top <= formCenter.Top Then
                tmrDown = False
                formOpened(2) = False
                formDown.Hide
                cmdShowDown.Enabled = True
            End If
        End If
    End With
End Sub

' Timer - Left
Private Sub tmrLeft_Timer()
    With formLeft
        If formOpened(3) = False Then
            .Left = .Left - lblStep
            If .Left <= formCenter.Left - formCenter.Width Then
                tmrLeft = False
                formOpened(3) = True
                cmdShowLeft.Enabled = True
            End If
        Else
            .Left = .Left + lblStep
            If .Left >= formCenter.Left Then
                tmrLeft = False
                formOpened(3) = False
                formLeft.Hide
                cmdShowLeft.Enabled = True
            End If
        End If
    End With

End Sub

' Timer - Right
Private Sub tmrRight_Timer()
    With formRight
        If formOpened(1) = False Then
            .Left = .Left + lblStep
            If .Left >= formCenter.Left + formCenter.Width Then
                tmrRight = False
                formOpened(1) = True
                cmdShowRight.Enabled = True
            End If
        Else
            .Left = .Left - lblStep
            If .Left <= formCenter.Left Then
                tmrRight = False
                formOpened(1) = False
                formRight.Hide
                cmdShowRight.Enabled = True
            End If
        End If
    End With
End Sub

' Timer - Up
Private Sub tmrUp_Timer()
    With formUp
        If formOpened(0) = False Then
            .Top = .Top - lblStep
            If .Top <= formCenter.Height - formCenter.Top Then
                tmrUp = False
                formOpened(0) = True
                cmdShowUp.Enabled = True
            End If
        Else
            .Top = .Top + lblStep
            If .Top >= formCenter.Height Then
                tmrUp = False
                formOpened(0) = False
                formUp.Hide
                cmdShowUp.Enabled = True
            End If
        End If
    End With
End Sub
