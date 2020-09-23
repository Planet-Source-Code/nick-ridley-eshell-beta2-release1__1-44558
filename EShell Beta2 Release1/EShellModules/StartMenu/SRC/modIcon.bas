Attribute VB_Name = "modIcon"
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

Const MAX_PATH = 260

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal flags&) As Long

Const SHGFI_DISPLAYNAME = &H200
Const SHGFI_EXETYPE = &H2000
Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Const SHGFI_LARGEICON = &H0        ' Large icon
Const SHGFI_SMALLICON = &H1        ' Small icon
Const ILD_TRANSPARENT = &H1        ' Display transparent
Const SHGFI_SHELLICONSIZE = &H4
Const SHGFI_TYPENAME = &H400
Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private shinfo As SHFILEINFO


Public Function DrawStartIcon(path As String, obj As Object, Optional small As Boolean = False, Optional index As Long = 0)
  
shinfo.iIcon = index
  
Dim hImgLarge&
  
hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
  
If Not small Then
    
    hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
    BASIC_SHGFI_FLAGS Or SHGFI_EXETYPE)

End If
  
obj.Cls
ImageList_Draw hImgLarge&, shinfo.iIcon, obj.hDC, 0, 0, ILD_TRANSPARENT
  
End Function


