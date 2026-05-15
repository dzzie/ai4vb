VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileSystem2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.7
'Site:         http://sandsprite.com
'License:      MIT


'Revision 3 <- Incompatiable with all previous..simplified & streamlined
'
'Info:     These are basically macros for VB's built in file processes
'            this should streamline your code quite a bit and hopefully
'            remove alot of redundant coding.
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie
    
'Changes Jan 5 05
'GetFreeFileName - fixed periodic overflow in
'GetFolderFiles - Added recursive option
'CreateFolder   - now returns boolean
'GetFreeFolderName - Added
'RandomNum - wrapped with 10 try error handling in case of periodic overflow
'Move - changed mechanism of copy to name x as y instead of copy delete
'CreateFile - now returns boolean
'
'changes feb 8 06
'   updated fileexists function to not throw err on bad path
'
' 5.21.12 - bugfix in RandomNum sporotic overflow time of day related...
' 11.14.13- bugfix in FileExists,FolderExists, GetFolderFiles recursive filter, GetParentFolder trailing \ bugfix
'  1.21.14- bugfix deletefile detects readonly attribute and removes
' 11.15.15- bugfix GetParentFolder could fail is ub folder name was repeated in path (sloppy replace)
' 5.31.17 = bugfix GetSubFolders.GetAttr could fail, FileExists could fail with embedded nulls (.\0file)

Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHFileOperation Lib "shell32" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type SHFILEOPSTRUCT
   hwnd        As Long
   wFunc       As Long
   pFrom       As String
   pTo         As String
   fFlags      As Integer
   fAborted    As Boolean
   hNameMaps   As Long
   sProgress   As String
 End Type
  
Private Const MAX_PATH As Long = 260
 
'File Operations
Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2
Private Const FO_DELETE As Long = &H3
Private Const FO_RENAME As Long = &H4

'File Operation Flags
Private Const FOF_MULTIDESTFILES As Long = &H1
Private Const FOF_CONFIRMMOUSE As Long = &H2
Private Const FOF_SILENT As Long = &H4              'don't create progress/report
Private Const FOF_RENAMEONCOLLISION As Long = &H8
Private Const FOF_NOCONFIRMATION As Long = &H10     'don't prompt the user.
Private Const FOF_WANTMAPPINGHANDLE As Long = &H20  'Fill in SHFILEOPSTRUCT.hNameMappings
                                                    'Must be freed using SHFreeNameMappings
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_FILESONLY As Long = &H80          'on *.*, do only files
Private Const FOF_SIMPLEPROGRESS As Long = &H100    'don't show names of files
Private Const FOF_NOCONFIRMMKDIR As Long = &H200    'don't confirm making any needed dirs
Private Const FOF_NOERRORUI As Long = &H400         'don't put up error UI
Private Const FOF_NORECURSION As Long = &H1000      'don't recurse into directories
Private Const FOF_NOCOPYSECURITYATTRIBS As Long = &H800  'don't copy NT file Security Attributes
Private Const FOF_NO_CONNECTED_ELEMENTS As Long = &H2000 'Shell v5+: don't operate on connected file elements
Private Const FOF_WANTNUKEWARNING As Long = &H4000       'Shell v5+: during delete operation, warn in nuking instead of recycling (partially overrides FOF_NOCONFIRMATION)
Private Const FOF_NORECURSEREPARSE As Long = &H8000&  'Shell v5.1+: treat reparse points as objects, not containers
                                                   


'if newFolderPath exists, copy source into it, else newFolderPath is the final path
'note this api can create multiple layers of folders in newPath if they dont already exist before the copy..
Public Function CopyFolder(ByVal fPath As String, newFolderPath As String) As Boolean

   Dim shf As SHFILEOPSTRUCT, FOF_FLAGS As Long, rv As Long, pf As String, finalPath As String

   If Not FolderExists(fPath) Then Exit Function
   
   If FolderExists(newFolderPath) Then 'they specified the new parent folder?
        finalPath = newFolderPath & "\" & Me.FileNameFromPath(fPath) 'replicate source folder name
        If FolderExists(finalPath) Then Exit Function 'trying to create itself?
   Else
        finalPath = newFolderPath 'they specified the complete new folder path
   End If
   
   FOF_FLAGS = FOF_SILENT Or FOF_NOCONFIRMATION Or FOF_NOERRORUI Or FOF_NOCONFIRMMKDIR
   fPath = fPath & vbNullChar

   With shf
     .wFunc = FO_COPY
     .pFrom = fPath
     .pTo = finalPath
     .fFlags = FOF_FLAGS
   End With

   rv = SHFileOperation(shf)
   If rv = 0 Then CopyFolder = True
   
End Function

Function GetFolderFiles(folderPath As String, Optional Filter As String = "*", Optional retFullPath As Boolean = True, Optional recursive As Boolean = False) As String()
   Dim fnames() As String
   Dim fs As String
   Dim folders() As String
   Dim i As Integer
   
   If Not FolderExists(folderPath) Then
        'returns empty array if fails
        GetFolderFiles = fnames()
        Exit Function
   End If
   
   folderPath = IIf(Right(folderPath, 1) = "\", folderPath, folderPath & "\")
   
   fs = Dir(folderPath & Filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then push fnames(), IIf(retFullPath = True, folderPath & fs, fs)
     fs = Dir()
   Wend
   
   If recursive Then
        folders() = GetSubFolders(folderPath)
        If Not AryIsEmpty(folders) Then
            For i = 0 To UBound(folders)
                FolderEngine folders(i), fnames(), Filter
            Next
        End If
        If Not retFullPath Then
            For i = 0 To UBound(fnames)
                fnames(i) = Replace(fnames(i), folderPath, Empty) 'make relative path from base
            Next
        End If
    End If
   
   GetFolderFiles = fnames()
End Function


Private Sub FolderEngine(fldrpath As String, ary() As String, Optional Filter As String = "*")

    Dim files() As String
    Dim folders() As String
    Dim i As Long
     
    files = GetFolderFiles(fldrpath, Filter)
    folders = GetSubFolders(fldrpath)
        
    If Not AryIsEmpty(files) Then
        For i = 0 To UBound(files)
            push ary, files(i)
        Next
    End If
    
    If Not AryIsEmpty(folders) Then
        For i = 0 To UBound(folders)
             FolderEngine folders(i), ary, Filter
        Next
    End If
    
End Sub

Function GetSubFolders(folderPath As String, Optional retFullPath As Boolean = True) As String()
    Dim fnames() As String
    Dim fd As String
    On Error Resume Next
    
    If Not FolderExists(folderPath) Then
        'returns empty array if fails
        GetSubFolders = fnames()
        Exit Function
    End If
    
   If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

   fd = Dir(folderPath, vbDirectory)
   While fd <> ""
     If Left(fd, 1) <> "." Then
        If (GetAttr(folderPath & fd) And vbDirectory) = vbDirectory Then
           If Err.Number = 0 Then
                push fnames(), IIf(retFullPath = True, folderPath & fd, fd)
           Else
                Err.Clear
           End If
        End If
     End If
     fd = Dir()
   Wend
   
   GetSubFolders = fnames()
End Function

Function FolderExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Asc(Left(tmp, 1)) = 0 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Asc(Left(path, 1)) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If InStr(path, Chr(0)) > 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function GetParentFolder(path) As String
    Dim tmp() As String
    Dim my_path
    Dim ub As String
    
    On Error GoTo hell
    If Len(path) = 0 Then Exit Function
    
    my_path = path
    While Len(my_path) > 0 And Right(my_path, 1) = "\"
        my_path = Mid(my_path, 1, Len(my_path) - 1)
    Wend
    
    tmp = Split(my_path, "\")
    tmp(UBound(tmp)) = Empty
    my_path = Replace(Join(tmp, "\"), "\\", "\")
    If VBA.Right(my_path, 1) = "\" Then my_path = Mid(my_path, 1, Len(my_path) - 1)
    
    GetParentFolder = my_path
    Exit Function
    
hell:
    GetParentFolder = Empty
    
End Function

Function CreateFolder(path As String) As Boolean
   On Error GoTo blah
   If FolderExists(path) Then Exit Function
   MkDir path
   If Not FolderExists(path) Then Exit Function
   CreateFolder = True
blah:
End Function

Function FileNameFromPath(FullPath As String) As String
    Dim tmp() As String
    If InStr(FullPath, "\") > 0 Then
        tmp = Split(FullPath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function WebFileNameFromPath(FullPath As String)
    Dim tmp() As String
    If InStr(FullPath, "/") > 0 Then
        tmp = Split(FullPath, "/")
        WebFileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function DeleteFile(fPath As String) As Boolean
 On Error GoTo hadErr
    
    Dim attributes As VbFileAttribute

    attributes = GetAttr(fPath)
    If (attributes And vbReadOnly) Then
        attributes = attributes - vbReadOnly
        SetAttr fPath, attributes
    End If

    Kill fPath
    DeleteFile = True
    
 Exit Function
hadErr:
'MsgBox "DeleteFile Failed" & vbCrLf & vbCrLf & fpath
DeleteFile = False
End Function

Sub Rename(FullPath As String, newName As String)
  Dim pf As String
  pf = GetParentFolder(FullPath)
  Name FullPath As pf & "\" & newName
End Sub

Sub SetAttribute(fPath, it As VbFileAttribute)
   SetAttr fPath, it
End Sub

Function GetExtension(path) As String
    Dim tmp() As String
    Dim ub As String
    If Len(path) = 0 Then Exit Function
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetExtension = LCase(Mid(ub, InStrRev(ub, "."), Len(ub)))
    Else
       GetExtension = ""
    End If
End Function

Function GetBaseName(path As String) As String
    Dim tmp() As String
    Dim ub As String
    If Len(path) = 0 Then Exit Function
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

'can also just accept a file name
Function ChangeExt(path As String, ext As String)
    Dim newPath As String
    
    If Left(ext, 1) <> "." Then ext = "." & ext
    newPath = GetBaseName(path) & ext
    
    If LCase(path) <> LCase(newPath) Then
        If FileExists(path) Then
            Rename path, newPath
        End If
    End If
    
    ChangeExt = newPath

End Function

Function SafeFileName(proposed As String) As String
  Dim badChars As String, bad() As String, i As Long
  badChars = ">,<,&,/,\,:,|,?,*,"""
  bad = Split(badChars, ",")
  For i = 0 To UBound(bad)
    proposed = Replace(proposed, bad(i), "")
  Next
  SafeFileName = CStr(proposed)
End Function

Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error Resume Next

    Do While 1
        Err.Clear
        Randomize
        tmp = Round(Timer * Now * Rnd(), 0)
        RandomNum = tmp
        If Err.Number = 0 Then Exit Function
        If tries < 100 Then
            tries = tries + 1
        Else
            Exit Do
        End If
    Loop
    
    RandomNum = GetTickCount
    
End Function

Function GetFreeFileName(ByVal folder As String, Optional extension = ".txt") As String
    
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String

    If Not FolderExists(folder) Then Exit Function
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
again:
    Do
      tmp = folder & RandomNum() & extension
    Loop Until Not FileExists(tmp)
    
    GetFreeFileName = tmp
    
Exit Function
handler:

    If i < 10 Then
        i = i + 1
        GoTo again
    End If
    
End Function


Function GetFreeFolderName(ByVal parentFolder As String, Optional prefix As String = "") As String
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String

    If Not FolderExists(parentFolder) Then Exit Function
    If Right(parentFolder, 1) <> "\" Then parentFolder = parentFolder & "\"
        
again:
    Do
      tmp = parentFolder & prefix & RandomNum()
    Loop Until Not FolderExists(tmp)
    
    GetFreeFolderName = tmp
    
Exit Function
handler:

    If i < 10 Then
        i = i + 1
        GoTo again
    End If
    
End Function



Function buildPath(folderPath As String) As Boolean
    On Error GoTo oops
    
    If FolderExists(folderPath) Then buildPath = True: Exit Function
    
    Dim tmp() As String, build As String, i As Long
    
    tmp = Split(folderPath, "\")
    build = tmp(0)
    For i = 1 To UBound(tmp)
        build = build & "\" & tmp(i)
        If InStr(tmp(i), ".") < 1 Then
            If Not FolderExists(build) Then CreateFolder (build)
        End If
    Next
    buildPath = True
    Exit Function
oops: buildPath = False
End Function


Function ReadFile(filename) As Variant
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub WriteFile(path As String, it As Variant)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Function WriteTempFile(ByVal content) As String
    Dim t As String
    t = GetFreeFileName(Environ("temp"))
    WriteFile t, content
    WriteTempFile = t
End Function

Sub AppendFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub


Function Copy(fPath As String, toFolder As String)
   Dim basename As String, newName As String
   If FolderExists(toFolder) Then
       basename = FileNameFromPath(fPath)
       toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
       newName = toFolder & basename
       FileCopy fPath, newName
       Copy = newName
   Else 'assume tofolder is actually new desired file path
       FileCopy fPath, toFolder
       Copy = toFolder
   End If
End Function

Public Function FileCopy2(ByVal pvSrc, ByVal pvTarg) As Boolean
    Dim fso
    On Error GoTo hell
    Set fso = CreateObject("Scripting.FileSystemObject")    '(reference: ms Scripting Runtime)
    fso.CopyFile pvSrc, pvTarg
    FileCopy2 = True
    Exit Function
hell:
End Function

Function Move(fPath As String, toFolder As String)
    Dim fname As String
    fname = FileNameFromPath(fPath)
    toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
    
    Name fPath As toFolder & fname
    Move = toFolder & fname
    
End Function

Function CreateFile(fPath As String) As Boolean
    On Error GoTo hell
    Dim f As Long
    f = FreeFile
    If FileExists(fPath) Then Exit Function
    Open fPath For Binary As f
    Close f
    If FileExists(fPath) Then CreateFile = True
hell:
End Function


Function DeleteFolder(folderPath As String, Optional Force As Boolean = True) As Boolean
 On Error GoTo failed
   Call delTree(folderPath, Force)
   RmDir folderPath
   DeleteFolder = True
 Exit Function
failed:  DeleteFolder = False
End Function

Private Sub delTree(folderPath As String, Optional Force As Boolean = True)
   Dim sfi() As String, sfo() As String, i As Integer
   sfi() = GetFolderFiles(folderPath)
   sfo() = GetSubFolders(folderPath)
   If Not AryIsEmpty(sfi) And Force = True Then
        For i = 0 To UBound(sfi)
            DeleteFile sfi(i)
        Next
   End If
   
   If Not AryIsEmpty(sfo) And Force = True Then
        For i = 0 To UBound(sfo)
            Call DeleteFolder(sfo(i), True)
        Next
   End If
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function FolderName(folderPath) As String
    Dim ret As String, tmp() As String
    If Len(folderPath) = 0 Then Exit Function
    tmp = Split(folderPath, "\")
    If Not AryIsEmpty(tmp) Then
        If Len(tmp(UBound(tmp))) <> 0 Then ret = tmp(UBound(tmp)) _
        Else ret = tmp(UBound(tmp) - 1)
    Else
        ret = CStr(folderPath)
    End If
    FolderName = ret
End Function


Private Sub Class_Initialize()
   ' If Not isRegistered And Not isInitalized Then TellThemAllAboutIt
End Sub
