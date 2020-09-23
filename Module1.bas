Attribute VB_Name = "Module1"
'declarations for LaunchWithDefaultApp
#If Win32 Then
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long


Private Declare Function GetDesktopWindow Lib "user32" () As Long
#Else


Declare Function ShellExecute Lib "SHELL" (ByVal hwnd%, _
    ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, _
    ByVal lpszDir$, ByVal fsShowCmd%) As Integer


Declare Function GetDesktopWindow Lib "USER" () As Integer
#End If
Private Const SW_SHOWNORMAL = 1
'End declaration for launchwithdefaultapp


Public Function LaunchWithDefaultApp(VData As String) As Long
      Dim Scr_hDC As Long
      Dim Dire As String
      Dire = Left(VData, 3)
      Scr_hDC = GetDesktopWindow()
      StartDoc = ShellExecute(Scr_hDC, "Open", VData, "", Dire, SW_SHOWNORMAL)
End Function
Private Function FileExists(ByVal PathToCheck As String) As Boolean
    'incovenients of the method:  Last access date and time will be changed to the current date and time
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open PathToCheck For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
               
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
       
    Exit Function
End Function
Private Function Rename(This As String, ToThis As String, Optional Descript As Boolean = False) As Boolean
    Select Case Descript
            Case False
               If FileExists(ToThis) Then
                   If MsgBox(ToThis & " already exixts, dou you want to verwrite the destination file ? ", vbApplicationModal + vbCritical + vbYesNo, "Warning , Destinationfile allready exixts ! ") = vbYes Then
                       Kill ToThis
                       Name This As ToThis
                   Else
                   MsgBox " Opreation Cancelled", vbInformation + vbOKOnly
                   End If
               Else
                   Name This As ToThis
               End If
            Case True
               Name This As ToThis
            End Select
End Function
Private Function MoveFile(This As String, There As String)
Move This, There
End Function

Public Function GetAttributes(OfThis As String) As String
Dim Tmp As VbFileAttribute
Tmp = GetAttr(OfThis)
If Tmp >= vbAlias Then '64
    GetAttributes = GetAttributes & " Alias"
    Tmp = Tmp - vbAlias
End If
If Tmp >= vbArchive Then ' 32
    GetAttributes = GetAttributes & " Archive"
    Tmp = Tmp - vbArchive
End If
If Tmp >= vbDirectory Then '16
    GetAttributes = GetAttributes & " directory"
    Tmp = Tmp - vbDirectory
End If
If Tmp >= vbVolume Then '8
    GetAttributes = GetAttributes & " volume"
    Tmp = Tmp - vbVolume
    End If

If Tmp >= vbSystem Then '4
    GetAttributes = GetAttributes & " System"
    Tmp = Tmp - vbSystem
End If
If Tmp >= vbHidden Then '2
    GetAttributes = GetAttributes & " Hidden"
    Tmp = Tmp - vbHidden
End If
If Tmp >= vbReadOnly Then '1
    GetAttributes = GetAttributes & " Read Only"
    Tmp = Tmp - vbReadOnly
End If
If Tmp = vbNormal Then '0
    GetAttributes = GetAttributes & " Normal"
    Tmp = Tmp - vbNormal
End If


End Function
Public Function SetAttributes(OfThis As String, ByVal ToThis As VbFileAttribute) As Boolean
 SetAttributes = True
 On Error GoTo errh
  SetAttr OfThis, ToThis
 GoTo fin
errh:
 SetAttributes = False
 Err.Clear
 Exit Function
fin:
 SetAttributes = True
End Function
Public Function DeleteFileAndDescr(ByVal FilePath As String, FileName As String) As Boolean
On Error Resume Next
Kill FilePath & "\" & FileName
Kill FilePath & "\descr\" & FileName & ".des"
End Function

