Attribute VB_Name = "FileAccess"
'***************************************
'* Special Note on use:                *
'* A reference must be made to the     *
'* Microsoft Scripting Library for     *
'* some of the functions to work.      *
'***************************************

'**************************************************************************************
'********************************** Private Declarations ******************************
'**************************************************************************************
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_CREATEPROGRESSDLG As Long = &H0
Public Enum RemoveMethod
    RecycleFile = 1
    DeleteFile = 2
End Enum

Private FileObject As New FileSystemObject
'***********************************************************************************
'***********************declarations for LaunchWithDefaultApp***********************
'***********************************************************************************
#If Win32 Then
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long


Private Declare Function GetDesktopWindow Lib "User32" () As Long
#Else


Declare Function ShellExecute Lib "SHELL" (ByVal hWnd%, _
    ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, _
    ByVal lpszDir$, ByVal fsShowCmd%) As Integer


Declare Function GetDesktopWindow Lib "USER" () As Integer
#End If
Private Const SW_SHOWNORMAL = 1
'End declaration for launchwithdefaultapp

'**************************************
'Windows API/Global Declarations for :Sy
'     stem Folders
'**************************************

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Const CSIDL_STARTMENU = &HB
Const CSIDL_DESKTOPDIRECTORY = &H10
Const CSIDL_DRIVES = &H11
Const CSIDL_NETWORK = &H12
Const CSIDL_NETHOOD = &H13
Const CSIDL_FONTS = &H14
Const CSIDL_TEMPLATES = &H15
Const MAX_PATH = 260
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

'*perso test
Public Type att
Alias As Integer
Archive As Integer
System As Integer
ReadOnly As Integer
Volume As Integer
Directory As Integer
Hidden As Integer
Normal As Integer
End Type
Global at As att
'********************************************************************

'******broseforfolder api declaration
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
'Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
'Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'******************************************************************************************************************************


'****************GetIconFromFile decaration**************************
Global Const DI_MASK = &H1
Global Const DI_IMAGE = &H2
Global Const DI_NORMAL = DI_MASK Or DI_IMAGE
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function DrawIconEx Lib "User32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "User32" (ByVal hIcon As Long) As Long
'********************************************************************
Public Function GetIconFromFile(filename As String) As Long
    GetIconFromFile = ExtractAssociatedIcon(App.hInstance, filename, 2)
End Function
Public Function BrowseFF() As String
    
    
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        'Set the owner window
        .hWndOwner = Form1.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("C:\", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    BrowseFF = sPath
End Function





'***********************************************************************************
'****************************CODE***************************************************
'***********************************************************************************




'*** Launch With Default App ********************
'* Iputs: VData - The name of the file          *
'* purpose: !!                                  *
'************************************************
Public Function LaunchWithDefaultApp(VData As String) As Long
      Dim Scr_hDC As Long
      Dim Dire As String
      Dire = Left(VData, 3)
      Scr_hDC = GetDesktopWindow()
      StartDoc = ShellExecute(Scr_hDC, "Open", VData, "", Dire, SW_SHOWNORMAL)
End Function



'*** File Exists 1 and 2*************************
'* Inputs: FileName - The name of the file.     *
'* Returns: True if the file exists.            *
'* Purpose: Tests whether or not a file exists. *
'************************************************
Public Function FileExists(filename As String) As Boolean
    FileExists = Not (Dir(filename) = "")
End Function

Private Function FileExists2(ByVal PathToCheck As String) As Boolean
    'incovenients of the method:  Last access date and time will be changed to the current date and time
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open PathToCheck For Input As #1
        Close #1
        'no error, file exists
        FileExists2 = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists2 = False
    Exit Function
End Function

'*** Rename File ***************************************
'* Inputs : This - file to be moved                    *
'*          Tothis - Location to be moved              *
'*          Description - working or not with...       *
'*******************************************************
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
'*** Get File Extension ********************************
'* Inputs: FileName - The name of the file.            *
'* Returns: The extension of the file as a string.     *
'* Purpose: Finds the extension of the specified file. *
'*******************************************************
Public Function GetFileExt(filename As String) As String
    Dim i As Integer
    i = Len(filename)
    While i > 1 And Mid$(filename, i, 1) <> "."
        i = i - 1
    Wend
    If Mid$(filename, i, 1) = "." Then
        GetFileExt = Right$(filename, Len(filename) - i)
    ElseIf i = 0 Then
        GetFileExt = ""
    End If
End Function


'*** Get File Name ************************************
'* Inputs: FileName - The name of the file.           *
'* Returns: The name of the file (without extension). *
'* Purpose: Finds the name of the specified file.     *
'******************************************************

Public Function GetFileName(filename As String) As String
    Dim i As Integer
    i = Len(filename)
    While i > 0 And Mid$(filename, i, 1) <> "\"
        i = i - 1
    Wend
    If Mid$(filename, i, 1) = "\" Then
        GetFileName = Right$(filename, Len(filename) - i)
    ElseIf i = 0 Then
        GetFileName = ""
    End If
End Function
Public Function GetFnameWITHOUText(filename As String) As String
Dim i As Integer
i = Len(filename)
While i > 1 And Mid$(filename, i, 1) <> "."
i = i - 1
Wend
If Mid$(filename, i, 1) = "." Then
    GetFnameWITHOUText = Left(filename, i - 1)
ElseIf i = 0 Then
    GetFnameWITHOUText = "Error"
End If


End Function

'*** Remove File ******************************
'* Inputs: FileName - The name of the file.   *
'*         Action - The method of deletion.   *
'* Purpose: Deletes a file or moves it to the *
'*          recycle bin.                      *
'* Notes: Possible values for "Action:"       *
'*        RecycleFile - Moves the file to the *
'*                      recycle bin.          *
'*        DeleteFile - Deletes the file.      *
'*        Both of these methods displays a    *
'*        confirmation prompt to the user.    *
'**********************************************
Public Function RemoveFile(filename As String, Action As RemoveMethod) As Boolean
    Dim FileOperation As SHFILEOPSTRUCT
    Dim tmpReturn As Long
    On Error GoTo RemoveFile_Err
    With FileOperation
        .wFunc = FO_DELETE
        .pFrom = filename
        If Action = RecycleFile Then
            .fFlags = FOF_ALLOWUNDO + FOF_CREATEPROGRESSDLG
        Else
            .fFlags = FO_DELETE + FOF_CREATEPROGRESSDLG
        End If
    End With
    tmpReturn = SHFileOperation(FileOperation)
    If tmpReturn <> 0 Then
        RemoveFile = False
    Else
        RemoveFile = True
    End If
    Exit Function
RemoveFile_Err:
    RemoveFile = False
End Function


'*** Create Directory ***************************
'* Inputs: DirName - The name of the directory. *
'* Purpose: Creates the specified directory.    *
'************************************************
Public Sub CreateDir(DirName As String)
On Error GoTo ErrH

    Call FileObject.CreateFolder(DirName)
    GoTo fin
ErrH:
If Err.Number = 58 Then Err.Clear Else MsgBox Err.Description
Err.Clear
fin:
End Sub


'*** Delete Directory *******************************
'* Inputs: DirName - The name of the directory.     *
'* Purpose: Deletes the specified directory.        *
'* Notes: It does not prompt for user confirmation. *
'*        The data within the directory is also     *
'*        deleted and not moved to the recycle bin. *
'****************************************************
Public Sub DeleteDir(DirName As String)
    Call FileObject.DeleteFolder(DirName, True)
End Sub


'*** Create Temporary File *********************
'* Inputs: none                                *
'* Returns: The name of the temporary file as  *
'*          a string.                          *
'* Purpose: Generates a unique name to be used *
'*          for a temporary file name.         *
'***********************************************
Public Function CreateTemp() As String
    CreateTemp = FileObject.GetTempName
End Function

'*** Get and Set attributes of o file ****************
'*                                                   *
'*****************************************************
Public Function GetAttributes(OfThis As String) As String
Dim Tmp As VbFileAttribute
Tmp = GetAttr(OfThis)
With at
.Alias = 0
.Archive = 0
.Directory = 0
.Hidden = 0
.Normal = 0
.ReadOnly = 0
.System = 0
.Volume = 0
End With

If Tmp >= vbAlias Then '64
    GetAttributes = GetAttributes & " Alias"
    Tmp = Tmp - vbAlias
    at.Alias = 1
End If
If Tmp >= vbArchive Then ' 32
    GetAttributes = GetAttributes & " Archive"
    Tmp = Tmp - vbArchive
    at.Archive = 1
    End If
If Tmp >= vbDirectory Then '16
    GetAttributes = GetAttributes & " directory"
    Tmp = Tmp - vbDirectory
    at.Directory = 1
End If
If Tmp >= vbVolume Then '8
    GetAttributes = GetAttributes & " volume"
    Tmp = Tmp - vbVolume
    at.Volume = 1
    End If

If Tmp >= vbSystem Then '4
    GetAttributes = GetAttributes & " System"
    Tmp = Tmp - vbSystem
    at.System = 1
End If
If Tmp >= vbHidden Then '2
    GetAttributes = GetAttributes & " Hidden"
    Tmp = Tmp - vbHidden
    at.Hidden = 1
End If
If Tmp >= vbReadOnly Then '1
    GetAttributes = GetAttributes & " Read Only"
    Tmp = Tmp - vbReadOnly
    at.ReadOnly = 1
End If
If Tmp = vbNormal Then '0
    GetAttributes = GetAttributes & " Normal"
    Tmp = Tmp - vbNormal
    at.Normal = 1
End If


End Function
Public Function SetAttributes(OfThis As String, ByVal ToThis As VbFileAttribute) As Boolean
 SetAttributes = True
 On Error GoTo ErrH
  SetAttr OfThis, ToThis
 GoTo fin
ErrH:
 SetAttributes = False
 Err.Clear
 Exit Function
fin:
 SetAttributes = True
End Function


'*** Write INI ***************************************
'* Inputs: SectionName - The name of the section     *
'*         to write to.                              *
'*         KeyName - The name of the key to write.   *
'*         KeyValue - The value to write to the key. *
'*         FileName - The name of the INI file to    *
'*         write to.                                 *
'* Purpose: Writes the specified value and name to   *
'*          an INI file.                             *
'* Notes: This function is included for Win16        *
'*        compatability only.  Whenever possible,    *
'*        data should be written to the              *
'*        registry instead.                          *
'*****************************************************
Public Sub WriteINI(SectionName As String, KeyName As String, KeyValue As String, filename As String)
    WritePrivateProfileString SectionName, KeyName, KeyValue, filename
End Sub


'*** Read INI *************************************
'* Inputs: SectionName - The name of the section  *
'*         from which to read.                    *
'*         KeyName - The name of the key whose    *
'*         value is to be read.                   *
'*         FileName - The name of the INI file.   *
'* Returns: The value of the specified key name.  *
'* Purpose: Reads a value from an INI file.       *
'* Notes: This function is included for Win16     *
'*        compatability only.  Whenever possible, *
'*        data should be written to the           *
'*        registry instead.                       *
'*        If the key was not found, the string    *
'*        "NOT FOUND" is returned.                *
'**************************************************
Public Function ReadINI(SectionName As String, KeyName As String, filename As String) As String
    Dim tmpBuffer As String * 255
    GetPrivateProfileString SectionName, KeyName, "NOT FOUND", tmpBuffer, Len(tmpBuffer), filename
    ReadINI = tmpBuffer
End Function


Private Function IsPathValid(ThisFile As String) As Long
    Dim lngPos As Long 'Declare container For postion of "\"
    Dim lngTemp As Long 'Declare a Temporary long Container For comparason
    Dim intLoop As Integer 'Declare container To work as counter
    Dim arrChars(21) As String 'Declare container For invalid charactors in path
    '---------------------------------------
    '     -----------------------------
    'Return Values And Descriptions
    '0 = No path set
    '1 = Valid Path set
    '2 = Invalid Charactor Found
    '3 = Invalid ":" found after second char
    '     of path
    '4 = Invalid "\\" found after second cha
    '     r of path
    '5 = No Drive assignment Found and Not a
    '     UNC Path
    '6 = Valid Drive letter found but no "\"
    '     preceding ":"
    '---------------------------------------
    '     -----------------------------
    'assign invalid charactors to array
    arrChars(0) = "'"
    arrChars(1) = """"
    arrChars(2) = "("
    arrChars(3) = ")"
    arrChars(4) = "!"
    arrChars(5) = "@"
    arrChars(6) = "#"
    arrChars(7) = "%"
    arrChars(8) = "^"
    arrChars(9) = "&"
    arrChars(10) = "*"
    arrChars(11) = "+"
    arrChars(12) = "="
    arrChars(13) = "<"
    arrChars(14) = ">"
    arrChars(15) = "?"
    arrChars(16) = "/"
    arrChars(17) = "."
    arrChars(18) = ","
    arrChars(19) = "`"
    arrChars(20) = ";"


    If Len(ThisFile) = 0 Then 'if length of Property is zero return an Error value
        IsPathValid = 0
        'clean out variables to conserve memory
        '     use
        GoTo ExitThisFunction
    End If


    For intLoop = 0 To 20 Step 1 'loop through array
        lngPos = InStr(1, ThisFile, arrChars(intLoop), vbBinaryCompare) 'Check For invalid charactor in String


        If lngPos > 0 Then 'if found then return Error value
            IsPathValid = 2
            'clean out variables to conserve memory
            '     use
            GoTo ExitThisFunction
        End If
    Next intLoop
    lngTemp = 1 'set Comparason starting point
    lngPos = InStr(1, ThisFile, ":", vbBinaryCompare) 'check For correct drive letter syntax


    Select Case lngPos
        Case Is = 1, Is >= 3 'invalid drive letter syntax found return Error value
        IsPathValid = 3
        'clean out variables to conserve memory
        '     use
        GoTo ExitThisFunction
        Case Is = 0 'No Drive letter assignment found check To see If unc path


        For intLoop = 1 To Len(ThisFile) Step 1 'set counter to step through Each charactor in String
            lngPos = InStr(intLoop, ThisFile, "\", vbBinaryCompare) 'check For directory delimiter


            If lngPos > 0 Then 'value found so check that it is Single


                If intLoop > 1 Then 'Starting values With "\\" are acceptable as unc paths


                    If lngPos > 2 And (lngPos - 1) = lngTemp Then 'any other location In the String With "\\" is invalid
                        IsPathValid = 4 'return Error value
                        'clean out variables to conserve memory
                        '     use
                        GoTo ExitThisFunction
                    End If
                End If
            Else


                If intLoop = 1 Or intLoop = 2 Then 'Must have at least "//" as first 2 charactors to be valid unc path
                    IsPathValid = 5 'return Error value
                    'clean out variables to conserve memory
                    '     use
                    GoTo ExitThisFunction
                End If
            End If
            lngTemp = lngPos 'increment temp value
        Next intLoop 'increment counter
        Case Is = 2
        'check rest of string for "\\"
        lngTemp = 1 'set Comparason starting point


        For intLoop = 3 To Len(ThisFile) Step 1 'set counter to step through Each charactor in String
            lngPos = InStr(intLoop, ThisFile, "\", vbBinaryCompare) 'check For directory delimiter


            If lngPos > 0 Then 'value found so check that it is Single


                If lngPos > 3 And (lngPos - 1) = lngTemp Then 'any other location In the String With "\\" is invalid
                    IsPathValid = 4 'return Error value
                    'clean out variables to conserve memory
                    '     use
                    GoTo ExitThisFunction
                End If
            Else


                If intLoop = 3 Then 'must have a "/" following drive letter To be valid path
                    IsPathValid = 6 'return Error value
                    'clean out variables to conserve memory
                    '     use
                    GoTo ExitThisFunction
                End If
            End If
            lngTemp = lngPos 'increment temp value
        Next intLoop 'increment counter
    End Select
IsPathValid = 1 'path passes checks return valid value
'clean out variables to conserve memory
'     use
ExitThisFunction:
lngPos = vbNull
intLoop = vbNull
lngTemp = vbNull
Erase arrChars
End Function


Public Function GetSpecialfolder(CSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    'Get the special folder
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = NOERROR Then
        'Create a buffer
        Path$ = SPACE$(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        'Remove the unnecessary chr$(0)'s
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Public Function GetFileSize(filename) As String
Dim s As Double
Dim st As String
s = FileLen(filename) / 1024
s = s / 1024
st = "Mb"
 If s > 1 Then
 GoTo ok
 Else
 s = s * 1024
 st = "Kb"
 End If
 If s > 1 Then
 GoTo ok
 Else
 s = s * 1024
 st = "b"
 End If
 
ok:
 GetFileSize = Round(s, 2) & st
fin:
End Function
Public Function GetFileDate(filename As String) As String
    On Error Resume Next
    GetFileDate = FileDateTime(filename)
End Function
Public Function PutInRecycle(ParamArray vntFileName() As Variant) As Boolean
   Dim i As Integer
   Dim sFileNames As String
   Dim SHFileOp As SHFILEOPSTRUCT

   For i = LBound(vntFileName) To UBound(vntFileName)
      sFileNames = sFileNames & vntFileName(i) & vbNullChar
   Next
        
   sFileNames = sFileNames & vbNullChar

   With SHFileOp
      .wFunc = FO_DELETE
      .pFrom = sFileNames
      .fFlags = FOF_ALLOWUNDO + FOF_SILENT + FOF_NOCONFIRMATION
   End With

   i = SHFileOperation(SHFileOp)
   
   If i = 0 Then
      PutInRecycle = True
   Else
      PutInRecycle = False
   End If
End Function
