VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1200
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Add personal comments"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox pc 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Dialog.frx":0000
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Click OK to automaticly generate description for"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub RichTextBox1_Change()

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then Me.Height = 3420 Else Me.Height = 1560
End Sub

Private Sub Form_Load()
Label1.Caption = Label1.Caption & Form1.File1.Path & "\" & Form1.File1.filename
End Sub

Private Sub OKButton_Click()
With Form1
.Des.TextRTF = _
    "File name : " & .Label1.Caption & "." & .Label2.Caption & Chr(13) & _
    "File Type : " & .Label3.Caption & Chr(13) & _
    "File Path : " & .File1.Path & Chr(13) & _
    "File Size : " & .Label4.Caption & Chr(13) & _
    "File Date : " & .Label5.Caption & Chr(13)
If Check1.Value = 1 Then .Des.TextRTF = .Des.TextRTF & pc.TextRTF
End With
Me.Visible = False
Unload Me

 
End Sub

