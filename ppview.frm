VERSION 5.00
Begin VB.Form ppview 
   Caption         =   "Form2"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   ScaleHeight     =   7500
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   1695
      Left            =   4920
      TabIndex        =   11
      Top             =   5640
      Width           =   1935
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Height          =   735
      Left            =   2280
      TabIndex        =   10
      Top             =   6600
      Width           =   2415
      Begin VB.Label status 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   855
      Left            =   2280
      TabIndex        =   7
      Top             =   5640
      Width           =   2415
      Begin VB.OptionButton Option6 
         Caption         =   "&Color"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         Caption         =   "&Black and withe"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Print"
      Height          =   975
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   960
      ScaleHeight     =   210
      ScaleMode       =   0  'User
      ScaleWidth      =   295
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "ppview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pdraft = -1
Const Plow = -2
Const Pmedium = -3
Const Phigh = -4
Dim Pq As Integer



Private Sub Command1_Click()
Picture1.Cls
Picture1.ScaleMode = 3
Picture1.PaintPicture Form1.ilist.ListImages(1).Picture, 0, 0
Picture1.CurrentY = Form1.ilist.ImageWidth '/ 100
Picture1.CurrentX = Form1.ilist.ImageHeight ' / 100
Picture1.Print " Tmx Directory tool ver " & App.Revision
Picture1.CurrentX = 0
Picture1.Line (0, Picture1.CurrentY)-(Picture1.Width, Picture1.CurrentY)
Picture1.Print " "
With Form1
Picture1.Print "Name : "; .Label1.Caption

Picture1.Print "File Type : "; .Label3.Caption
Picture1.Print "Size : "; .Label4.Caption
Picture1.Print "Creation Date : "; .Label5.Caption
End With
Picture1.Line (0, Picture1.CurrentY)-(Picture1.Width, Picture1.CurrentY)
Picture1.Print " "
Picture1.Print Form1.Des.Text
End Sub

Public Sub displayp(filename As String)


End Sub

Private Sub Command2_Click()
'Printer.Cls
status.Caption = "Now Printing on " & Printer.DeviceName
Dim t As ScaleModeConstants
t = vbMillimeters
If Option1.Value = True Then Pq = Pdraft
If Option2.Value = True Then Pq = Plow
If Option3.Value = True Then Pq = Pmedium
If Option4.Value = True Then Pq = Phigh
If Option5.Value = True Then Printer.ColorMode = vbPRCMMonochrome
If Option6.Value = True Then Printer.ColorMode = vbPRCMColor
On Error GoTo ErrHan
'GoTo Testt
Printer.PrintQuality = Pq
Printer.Print Tab(30)
Printer.ScaleMode = t
Printer.PaintPicture Form1.ilist.ListImages(1).Picture, 0, 0
Printer.Print " "
Printer.Print " "
Printer.CurrentX = 14
Printer.Print " Tmx Directory tool ver " & App.Revision
Printer.CurrentX = 0
Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
Printer.Print Tab(30)
Printer.Print " "
With Form1

Printer.Print Tab(30); "Name : "; .Label1.Caption
Printer.Print Tab(30); 30; "File Type : "; .Label3.Caption; "("; LCase(.Label2.Caption); ")"
Printer.Print Tab(30); "Size : "; .Label4.Caption
Printer.Print Tab(30); "Creation Date : "; .Label5.Caption
End With
Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
Printer.Print Tab(30)
Printer.Print " "
Printer.Print Form1.Des.Text
Printer.EndDoc
Testt:

status.Caption = Printer.DeviceName & " impression pool updated with new job"
GoTo fin

ErrHan:

status.Caption = "error, job calcelled"


fin:

End Sub

Private Sub Form_Load()

Timer1.Enabled = True
Command2.Picture = Form1.iconList.ListImages(32).Picture
Option1.Value = True
Option6.Value = True
Frame1.Caption = "Print quality"
Option1.Caption = "Draft"
Option2.Caption = "Low"
Option3.Caption = "Medium"
Option4.Caption = "High"
ListPrinters
End Sub

Private Sub Timer1_Timer()
Command1_Click
'testptr
Timer1.Enabled = False
End Sub

Private Sub ListPrinters()
Combo1.Clear
Dim x As Printer
   For Each x In Printers
      Combo1.AddItem Printer.DeviceName
   Next
Set x = Nothing
End Sub
