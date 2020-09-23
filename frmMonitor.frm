VERSION 5.00
Begin VB.Form frmMonitor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CPU Mon"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   64
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00008000&
      Height          =   975
      Left            =   1080
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   0
      Width           =   3615
   End
   Begin VB.PictureBox picUsage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   0
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   0
      Top             =   0
      Width           =   990
      Begin VB.Label lblCpuUsage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmMonitor - copyright © 2001, The KPD-Team
'Visit our site at http://www.allapi.net
'or email us at KPDTeam@allapi.net
Option Explicit
Private QueryObject As Object
Private Sub Form_Activate()
    'set this form always on top
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Me.Top = 0
Me.Left = 0
End Sub
Private Sub Form_Load()
    'set the Priority of this process to 'High'
    'this makes sure our program gets updated, even when
    'another process is consuming lots of CPU cycles
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    'Initialize our Query object
    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
    Else
        Set QueryObject = New clsCPUUsage
    End If
    'Initializing is necesarry for the correct values to be retrieved
    QueryObject.Initialize
    'start the timer
    tmrRefresh.Enabled = True
    'don't wait for the first interval to elapse
    tmrRefresh_Timer
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'stop the timer
    tmrRefresh.Enabled = False
    'clean up
    QueryObject.Terminate
    Set QueryObject = Nothing
    Form1.Tool.Buttons("cpu").MixedState = False
    Form1.Tool.Buttons("cpu").ToolTipText = " Load Cpu Monitor"
End Sub
Private Sub tmrRefresh_Timer()
    Dim Ret As Long
    'query the CPU usage
    Ret = QueryObject.Query
    If Ret = -1 Then
        tmrRefresh.Enabled = False
        lblCpuUsage.Caption = ":("
        MsgBox "Error while retrieving CPU usage"
    Else
        DrawUsage Ret, picUsage, picGraph
        lblCpuUsage.Caption = CStr(Ret) + "%"
    End If
End Sub
