VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stopwatch"
   ClientHeight    =   4335
   ClientLeft      =   1065
   ClientTop       =   7215
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   14760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7440
      TabIndex        =   3
      Top             =   2040
      Width           =   7215
   End
   Begin VB.CommandButton btnStartStop 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14160
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   0
   End
   Begin VB.Label lblmils1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   11880
      TabIndex        =   13
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Label lblmils2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   10680
      TabIndex        =   12
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   9720
      TabIndex        =   11
      Top             =   -240
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   6600
      TabIndex        =   10
      Top             =   -240
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3240
      TabIndex        =   9
      Top             =   -240
      Width           =   1335
   End
   Begin VB.Label lblhou2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   960
      TabIndex        =   8
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Label lblhou1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2160
      TabIndex        =   7
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Label lblmin2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4320
      TabIndex        =   6
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Label lblmin1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5520
      TabIndex        =   5
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Label lblsec2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   7560
      TabIndex        =   4
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Label lblCurTime 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lblsec1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8760
      TabIndex        =   0
      Top             =   -120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim timeArr(0 To 3, 0 To 1) As Integer      'Timer digits Array
Dim startStop As Boolean                    'StartStop Flip Variable

'Reset the stopwatch!
'Note* : this button is multi-state, and therefore makes a descision based on
'        whether the stopwatch is running.
Private Sub btnReset_Click()
    If startStop = False Then       'Reset all values if the timer is no longer running
        lblmils1.Caption = 0
        lblmils2.Caption = 0
        lblsec1.Caption = 0
        lblsec2.Caption = 0
        lblmin1.Caption = 0
        lblmin2.Caption = 0
        lblhou1.Caption = 0
        lblhou2.Caption = 0
        timeArr(0, 0) = 0   'Mils1
        timeArr(0, 1) = 0   'Mils2
        timeArr(1, 0) = 0   'Sec1
        timeArr(1, 1) = 0   'Sec2
        timeArr(2, 0) = 0   'Min1
        timeArr(2, 1) = 0   'Min2
        timeArr(3, 0) = 0   'Hou1
        timeArr(3, 1) = 0   'Hou2
        lapnum = 0
        frmLap.txtLap.Text = ""
    Else                            'Write a line to the lap form if the timer is running
        frmLap.Visible = True
        lapnum = lapnum + 1
        timelap = timeArr(3, 1) & timeArr(3, 0) & ":" & timeArr(2, 1) & timeArr(2, 0) & ":" & timeArr(1, 1) & timeArr(1, 0) & "." & timeArr(0, 1) & timeArr(0, 0)
        recordLap      'function is located in the module
    End If
End Sub

'Start/stop the stopwatch
Private Sub btnStartStop_Click()
    If startStop = True Then            'StartStop Flip-flop to change whether the timer is enabled or disabled
        startStop = False
        tmrMain.Enabled = False
        btnStartStop.Caption = "Start"
        btnReset.Caption = "Reset"
    ElseIf startStop = False Then
        startStop = True
        tmrMain.Enabled = True
        btnStartStop.Caption = "Stop"
        btnReset.Caption = "Lap"
    End If
End Sub

'End the program
Private Sub cmdClose_Click()
    End
End Sub

'Timer tick function for updating the clock
Private Sub Timer1_Timer()
    lblCurTime.Caption = Format(Now, "hh:mm:ss AMPM")
End Sub

'timer tick function for updating the stopwatch
Private Sub tmrMain_Timer()
    If timeArr(0, 0) = 9 Then
        timeArr(0, 0) = 0
        If timeArr(0, 1) = 5 Then
            timeArr(0, 1) = 0
            If timeArr(1, 0) = 9 Then
                timeArr(1, 0) = 0
                If timeArr(1, 1) = 5 Then
                    timeArr(1, 1) = 0
                    If timeArr(2, 0) = 9 Then
                        timeArr(2, 0) = 0
                        If timeArr(2, 1) = 5 Then
                            timeArr(2, 1) = 0
                            If timeArr(3, 0) = 9 Then
                                timeArr(3, 0) = 0
                                If timeArr(3, 1) = 5 Then
                                    timeArr(3, 1) = 0
                                Else
                                    timeArr(3, 1) = timeArr(3, 1) + 1
                                End If
                            Else
                                timeArr(3, 0) = timeArr(3, 0) + 1
                            End If
                        Else
                            timeArr(2, 1) = timeArr(2, 1) + 1
                        End If
                    Else
                        timeArr(2, 0) = timeArr(2, 0) + 1
                    End If
                Else
                    timeArr(1, 1) = timeArr(1, 1) + 1
                End If
            Else
                timeArr(1, 0) = timeArr(1, 0) + 1
            End If
        Else
            timeArr(0, 1) = timeArr(0, 1) + 1
        End If
    Else
        timeArr(0, 0) = timeArr(0, 0) + 1
    End If
    'print all values into textboxes
    lblmils1.Caption = timeArr(0, 0)
    lblmils2.Caption = timeArr(0, 1)
    lblsec1.Caption = timeArr(1, 0)
    lblsec2.Caption = timeArr(1, 1)
    lblmin1.Caption = timeArr(2, 0)
    lblmin2.Caption = timeArr(2, 1)
    lblhou1.Caption = timeArr(3, 0)
    lblhou2.Caption = timeArr(3, 1)
End Sub
