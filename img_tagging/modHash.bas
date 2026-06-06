Attribute VB_Name = "modHash"
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.8
'License:      MIT
'
' File hashing via the Windows CryptoAPI (advapi32). CFileSystem2 has no
' hashing, so this fills the gap. MD5 by default; to switch to SHA-256,
' change CALG_MD5 -> CALG_SHA_256 and the 16-byte buffer -> 32.

Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, _
    ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As Long, ByRef pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
    ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Byte, _
    ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Private Const PROV_RSA_AES      As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_MD5          As Long = &H8003&
Private Const CALG_SHA_256      As Long = &H800C&
Private Const HP_HASHVAL        As Long = 2

Private Type BROWSEINFO
    hOwner         As Long
    pidlRoot       As Long
    pszDisplayName As String
    lpszTitle      As String
    ulFlags        As Long
    lpfn           As Long
    lParam         As Long
    iImage         As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBI As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_EDITBOX          As Long = &H10
Private Const BIF_NEWDIALOGSTYLE   As Long = &H40   ' resizable + "Make New Folder"
Private Const MAX_PATH             As Long = 260

Public Function BrowseForFolder(ByVal hOwner As Long, Optional ByVal title As String = "Select a folder") As String
    Dim bi As BROWSEINFO, pidl As Long, path As String, n As Long

    bi.hOwner = hOwner
    bi.lpszTitle = title
    bi.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE Or BIF_EDITBOX

    pidl = SHBrowseForFolder(bi)
    If pidl = 0 Then Exit Function          ' user cancelled

    path = String$(MAX_PATH, vbNullChar)
    If SHGetPathFromIDList(pidl, path) <> 0 Then
        n = InStr(path, vbNullChar)
        If n > 0 Then path = Left$(path, n - 1)
        BrowseForFolder = path
    End If
    CoTaskMemFree pidl                      ' free the shell-allocated PIDL
End Function

' Returns lowercase hex MD5 of the file's contents, or "" on failure.
Public Function FileMD5(ByVal path As String) As String
    Dim hProv As Long, hHash As Long
    Dim b() As Byte, hashv(0 To 15) As Byte
    Dim hlen As Long, i As Long, s As String, f As Long, n As Long

    On Error GoTo done

    n = fileLen(path)
    If n > 0 Then
        ReDim b(0 To n - 1)
        f = FreeFile
        Open path For Binary Access Read As #f
        Get #f, , b
        Close #f
    End If

    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then Exit Function
    If CryptCreateHash(hProv, CALG_MD5, 0, 0, hHash) = 0 Then GoTo cleanup

    If n > 0 Then CryptHashData hHash, b(0), n, 0   ' empty file -> hash of nothing (d41d8cd9...)

    hlen = 16
    If CryptGetHashParam(hHash, HP_HASHVAL, hashv(0), hlen, 0) = 0 Then GoTo cleanup

    For i = 0 To 15
        s = s & Right$("0" & Hex$(hashv(i)), 2)
    Next
    FileMD5 = LCase$(s)

cleanup:
    If hHash <> 0 Then CryptDestroyHash hHash
    If hProv <> 0 Then CryptReleaseContext hProv, 0
done:
End Function
